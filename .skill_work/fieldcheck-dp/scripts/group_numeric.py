"""
group_numeric.py — Preview range/bucket groups for a NUM question's raw
values, before deciding how to configure its stub entry.

Two modes, reading the actual value distribution from rawdata.csv (never a
guess):

  Equal-width (default) — snaps bin width to a "nice" round number
  (1/2/5/10/20/25/50/100 * 10^n — the same rule chart axes use for tick
  spacing). Good for reasonably even distributions (age, counts). Prints a
  ready-to-paste static "choices" array (group + hidden entries).

  --quantile N — splits into N bins with ~equal respondent counts each
  (pandas qcut). Good for SKEWED data (income, spending) where equal-width
  bins would cram most respondents into 1-2 buckets. This mode is applied
  LIVE by the pipeline itself (table_generator.py) via a single stub config
  key — "num_quantile": N — so nothing needs to be pasted; this preview
  just lets you see the resulting bins/sizes before picking N.

Usage:
  python group_numeric.py <rawdata.csv> <metadata.json> --question S3_1 [--bins 6] [--width 5]
  python group_numeric.py <rawdata.csv> <metadata.json> --question D3 --quantile 5

Examples:
  python group_numeric.py output/VN8996/data/rawdata.csv output/VN8996/data/metadata.json --question S3_1
  python group_numeric.py output/VN8996/data/rawdata.csv output/VN8996/data/metadata.json --question S3_1 --width 10
  python group_numeric.py output/VN8996/data/rawdata.csv output/VN8996/data/metadata.json --question D3 --quantile 5
"""

import argparse
import json
import math
import sys
from pathlib import Path

import pandas as pd


def _nice_number(x: float, round_up: bool) -> float:
    """Snap x to the nearest "nice" 1/2/5/10 * 10^n (Heckbert's algorithm —
    the same rule used to pick axis tick spacing on charts)."""
    if x <= 0:
        return 1.0
    exp = math.floor(math.log10(x))
    frac = x / (10 ** exp)
    if round_up:
        nice_frac = 1 if frac <= 1 else 2 if frac <= 2 else 5 if frac <= 5 else 10
    else:
        nice_frac = 1 if frac < 1.5 else 2 if frac < 3 else 5 if frac < 7 else 10
    return nice_frac * (10 ** exp)


def propose_bin_width(lo: float, hi: float, target_bins: int) -> float:
    if hi <= lo:
        return 1.0
    return _nice_number((hi - lo) / target_bins, round_up=True)


def build_bins(values: list[float], width: float) -> list[tuple[float, float]]:
    """[start, start+width) bins covering all values, start aligned to a
    multiple of width below the minimum value."""
    lo, hi = min(values), max(values)
    start = math.floor(lo / width) * width
    edges = []
    cur = start
    while cur <= hi:
        edges.append((cur, cur + width))
        cur += width
    return edges


def _fmt_num(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else str(v)


def _bin_label(lo: float, hi: float, is_int: bool, is_last: bool) -> str:
    if is_last:
        return f"{_fmt_num(lo)}+"
    if is_int:
        return f"{_fmt_num(lo)}-{_fmt_num(hi - 1)}"
    return f"{_fmt_num(lo)}-{_fmt_num(hi)}"


def propose_groups(values: list[float], width: float | None, target_bins: int) -> list[dict]:
    """Return [{"label", "lo", "hi", "values": [...]}] — one dict per bin
    that actually contains at least one real value (empty bins dropped)."""
    is_int = all(float(v).is_integer() for v in values)
    lo, hi = min(values), max(values)
    if width is None:
        width = propose_bin_width(lo, hi, target_bins)
    edges = build_bins(values, width)
    result = []
    for i, (b_lo, b_hi) in enumerate(edges):
        is_last = i == len(edges) - 1
        members = sorted(v for v in set(values) if b_lo <= v < b_hi)
        if not members:
            continue
        result.append({
            "label": _bin_label(b_lo, b_hi, is_int, is_last),
            "lo": b_lo, "hi": b_hi,
            "values": members,
        })
    return result


def print_summary(question: str, values: list[float], groups: list[dict]) -> None:
    total = len(values)
    print(f"\n{question} — {total} respondents, range {min(values):g}-{max(values):g}")
    print("\nProposed groups:")
    for g in groups:
        cnt = sum(1 for v in values if v in g["values"])
        pct = cnt / total * 100 if total else 0
        print(f"  {g['label']:>10}  n={cnt:<5} ({pct:5.1f}%)  values={[_fmt_num(v) for v in g['values']]}")


def print_json_snippet(groups: list[dict]) -> None:
    print('\nReady-to-paste "choices" for datatable.json:')
    lines = ['  "choices": [']
    for g in groups:
        codes = ", ".join(f'"{_fmt_num(v)}"' for v in g["values"])
        lines.append(f'    {{ "codes": [{codes}], "label": "{g["label"]}" }},')
    for g in groups:
        for v in g["values"]:
            lines.append(f'    {{ "code": "{_fmt_num(v)}", "hidden": true }},')
    if lines[-1].endswith(","):
        lines[-1] = lines[-1][:-1]
    lines.append("  ]")
    print("\n".join(lines))


def propose_quantile_groups(values: list[float], n_quantiles: int) -> list[dict]:
    """~Equal-count bins (unlike propose_groups' equal-width) — mirrors
    table_generator.py's _quantile_num_groups exactly, so this preview
    matches what "num_quantile": N actually produces at table time.

    Labels use each bin's own real min/max member (inclusive on both ends —
    unlike propose_groups' equal-width bins, there's no separate exclusive
    upper edge here, since the cut points come from the data itself)."""
    import pandas as pd
    s = pd.Series(values)
    is_int = bool((s == s.round()).all())
    try:
        cats = pd.qcut(s, q=n_quantiles, duplicates="drop")
    except ValueError:
        return []
    intervals = sorted(cats.cat.categories, key=lambda iv: iv.left)
    result = []
    for i, interval in enumerate(intervals):
        members = sorted(s[cats == interval].unique())
        if not members:
            continue
        lo, hi = members[0], members[-1]
        is_last = i == len(intervals) - 1
        label = f"{_fmt_num(lo)}+" if is_last else f"{_fmt_num(lo)}-{_fmt_num(hi)}"
        result.append({"label": label, "values": members})
    return result


def print_quantile_config(question: str, n_quantiles: int) -> None:
    print(f'\nReady-to-use stub config (pipeline computes bins live, nothing to paste):')
    print(f'  {{ "question": "{question}", "stats": ["base", "percent"], "num_quantile": {n_quantiles} }}')


def main() -> None:
    parser = argparse.ArgumentParser(description="Propose range groups for a NUM question")
    parser.add_argument("rawdata", help="Path to rawdata.csv")
    parser.add_argument("metadata", help="Path to metadata.json")
    parser.add_argument("--question", required=True, help="Question label (e.g. S3_1)")
    parser.add_argument("--bins", type=int, default=6, help="Target number of bins (default 6)")
    parser.add_argument("--width", type=float, default=None,
                         help="Force a specific bin width instead of auto-picking one")
    parser.add_argument("--quantile", type=int, default=None,
                         help="Use ~equal-count quantile bins instead of equal-width "
                              "(good for skewed data like income/spending); overrides --bins/--width")
    args = parser.parse_args()

    rawdata_path = Path(args.rawdata)
    meta_path = Path(args.metadata)
    for p in (rawdata_path, meta_path):
        if not p.exists():
            print(f"ERROR: File not found: {p}", file=sys.stderr)
            sys.exit(1)

    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    questions = meta.get("questions", {})
    questions = questions.values() if isinstance(questions, dict) else questions
    q = next((q for q in questions if q.get("label") == args.question), None)
    if q is None:
        print(f"ERROR: Question '{args.question}' not found in metadata.", file=sys.stderr)
        sys.exit(1)
    if q.get("answer_type") != "NUM":
        print(f"ERROR: '{args.question}' is answer_type={q.get('answer_type')!r}, not NUM.",
              file=sys.stderr)
        sys.exit(1)

    df = pd.read_csv(rawdata_path, encoding="utf-8-sig", low_memory=False)
    if args.question not in df.columns:
        print(f"ERROR: Column '{args.question}' not found in rawdata.csv.", file=sys.stderr)
        sys.exit(1)

    values = pd.to_numeric(df[args.question], errors="coerce").dropna().tolist()
    if not values:
        print(f"No numeric values found for '{args.question}'.", file=sys.stderr)
        sys.exit(1)

    if args.quantile:
        groups = propose_quantile_groups(values, args.quantile)
        print_summary(args.question, values, groups)
        print_quantile_config(args.question, args.quantile)
    else:
        groups = propose_groups(values, args.width, args.bins)
        print_summary(args.question, values, groups)
        print_json_snippet(groups)


if __name__ == "__main__":
    main()
