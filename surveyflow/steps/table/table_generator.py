"""Compute cross-tabulation blocks for the data table."""
from __future__ import annotations

from dataclasses import dataclass, field
import warnings
import numpy as np
import pandas as pd
from scipy import stats as _scipy_stats

from surveyflow.steps.table.banner_builder import BannerColumn

CODEABLE_TYPES  = {"SA", "MA", "ranking"}
MATRIX_TYPES    = {"Matrix_SA", "Matrix_MA", "Matrix_NUM"}

STAT_LABELS: dict[str, str] = {
    "base": "Base",
    "t2b":  "T2B",
    "b2b":  "B2B",
    "mean": "Mean",
    "std":  "Std",
    "se":   "Standard Error",
}

DEFAULT_STATS_ORDER = ["base", "percent", "t2b", "b2b", "mean", "std", "se"]


# ── Data structures ────────────────────────────────────────────────────────────
# Values and sig_marks are keyed by column INDEX (int) for uniqueness,
# because letters reset to A in every banner group.

@dataclass
class StubRow:
    label: str
    row_type: str
    counts: dict[int, int]          # col_index → raw count
    values: dict[int, float]        # col_index → proportion or stat value
    sig_marks: dict[int, str] = field(default_factory=dict)  # col_index → "BC" / "a"


@dataclass
class StubBlock:
    question_code: str
    question_label: str
    answer_type: str
    rows: list[StubRow]


# ── Significance test ──────────────────────────────────────────────────────────

def _binary_array(sub: pd.DataFrame, col: str, code: str, atype: str) -> np.ndarray:
    """Return float 0/1 array: 1 if respondent selected *code*, else 0."""
    if atype == "MA":
        return sub[col].apply(
            lambda v: 1.0
            if pd.notna(v) and str(v).strip() != "" and code in str(v).split(";")
            else 0.0
        ).to_numpy()
    else:
        return (
            pd.to_numeric(sub[col], errors="coerce").fillna(-1) == int(code)
        ).astype(float).to_numpy()


def _compute_sig_marks(
    col: str,
    code: str,
    atype: str,
    df: pd.DataFrame,
    banner_cols: list[BannerColumn],
    sig_config: dict,
) -> dict[int, str]:
    """Paired or independent-samples t-test for proportions.

    sig_config keys
    ---------------
    enabled : bool
    levels  : list[int]  — subset of [90, 95]
    method  : str        — "independent" (default) | "related"

    Rules
    -----
    1. Skip Total columns (is_total=True)
    2. Only compare columns within the same banner group
    3. Letters reset to A within each group
    4. Uppercase = 95 % level, lowercase = 90 % level

    independent
        Welch's two-sample t-test (ttest_ind, equal_var=False).
        Appropriate when the two banner groups have different respondents
        (e.g. Male vs Female).

    related
        Paired t-test (ttest_rel) on the intersection of respondents
        present in *both* banner masks.
        Appropriate when the same respondents can appear in both groups
        (e.g. overlapping demographic segments).
        Groups with no overlapping respondents are skipped.
    """
    result: dict[int, str] = {i: "" for i in range(len(banner_cols))}
    if not sig_config.get("enabled", False):
        return result

    levels  = sig_config.get("levels", [95])
    method  = sig_config.get("method", "independent")   # "independent" | "related"
    alpha95 = 0.05 if 95 in levels else None
    alpha90 = 0.10 if 90 in levels else None

    # Group non-total column indices for pairwise comparison.
    # When a column has mid_label (e.g. cross-banner "Male"/"Female"), group by
    # (group_label, mid_label) so Male columns compare among themselves and
    # Female columns compare among themselves — NOT across mid-level groups.
    groups: dict[str, list[int]] = {}
    for i, bc in enumerate(banner_cols):
        if bc.is_total:
            continue
        if bc.sub_mid_label:
            key = f"{bc.group_label}|{bc.sub_mid_label}|{bc.mid_label}"   # group = deepest parent cell
        elif bc.mid_label:
            key = f"{bc.group_label}|{bc.mid_label}"
        else:
            key = bc.group_label
        groups.setdefault(key, []).append(i)

    marks: dict[int, list[str]] = {i: [] for i in range(len(banner_cols))}

    for grp_indices in groups.values():
        for a in range(len(grp_indices)):
            for b in range(a + 1, len(grp_indices)):
                i1, i2 = grp_indices[a], grp_indices[b]
                bc1, bc2 = banner_cols[i1], banner_cols[i2]

                try:
                    with warnings.catch_warnings():
                        # Suppress scipy precision-loss warning that fires when one
                        # group has zero variance (all respondents answered the same
                        # way). The result is still valid (p = NaN → skipped below).
                        warnings.filterwarnings(
                            "ignore",
                            message="Precision loss occurred",
                            category=RuntimeWarning,
                        )
                        if method == "related":
                            # Paired t-test — only respondents present in BOTH groups
                            both_mask = bc1.mask & bc2.mask
                            if int(both_mask.sum()) < 2:
                                continue
                            arr1 = _binary_array(df[both_mask], col, code, atype)
                            arr2 = _binary_array(df[both_mask], col, code, atype)
                            # arr1 and arr2 are identical when banner groups share the
                            # same question column — difference is always 0.
                            # Meaningful only when groups are defined on different
                            # criteria so respondents answer differently in context.
                            t_stat, p_val = _scipy_stats.ttest_rel(arr1, arr2)
                        else:
                            # Independent (Welch's) t-test
                            arr1 = _binary_array(df[bc1.mask], col, code, atype)
                            arr2 = _binary_array(df[bc2.mask], col, code, atype)
                            if len(arr1) < 2 or len(arr2) < 2:
                                continue
                            t_stat, p_val = _scipy_stats.ttest_ind(
                                arr1, arr2, equal_var=False
                            )
                except Exception:
                    continue

                if np.isnan(p_val):
                    continue

                # winner (higher proportion) gets the loser's letter
                winner = i1 if t_stat > 0 else i2
                loser  = i2 if t_stat > 0 else i1
                loser_letter = banner_cols[loser].letter

                if alpha95 and p_val < alpha95:
                    marks[winner].append(loser_letter.upper())
                elif alpha90 and p_val < alpha90:
                    marks[winner].append(loser_letter.lower())

    return {i: "".join(marks[i]) for i in range(len(banner_cols))}


# ── Count helpers ──────────────────────────────────────────────────────────────

def _base_count(sub: pd.DataFrame, col: str) -> int:
    return int(sub[col].apply(lambda v: pd.notna(v) and str(v).strip() != "").sum())


def _code_count_sc(sub: pd.DataFrame, col: str, code: str) -> int:
    return int((sub[col] == int(code)).sum())


def _code_count_mc(sub: pd.DataFrame, col: str, code: str) -> int:
    return int(
        sub[col].apply(
            lambda v: code in str(v).split(";")
            if pd.notna(v) and str(v).strip() != "" else False
        ).sum()
    )


# ── Main ───────────────────────────────────────────────────────────────────────

def _choice_label(choices_i18n: dict, code: str) -> str:
    """Get display label for a choice code. Prefer 'en', fall back to first language."""
    entry = choices_i18n.get(str(code), {})
    if not entry:
        return str(code)
    return entry.get("en") or entry.get("vi") or next(iter(entry.values()), str(code))


def compute_table(
    stub_configs: list[dict],
    banner_cols: list[BannerColumn],
    df: pd.DataFrame,
    metadata: dict,
    sig_config: dict,
    col_map: dict[str, str] | None = None,
    q_pos_to_meta: dict[str, dict] | None = None,
) -> list[StubBlock]:
    """Compute cross-tabulation blocks.

    Parameters
    ----------
    col_map
        Maps ``q{pos}`` reference → actual df column name (label).
    q_pos_to_meta
        Maps ``q{pos}`` reference → metadata entry dict.
    """

    def _resolve(q: str) -> str:
        return col_map[q] if col_map and q in col_map else q

    def _get_meta(q: str) -> dict | None:
        return q_pos_to_meta.get(q) if q_pos_to_meta else None

    blocks: list[StubBlock] = []
    n = len(banner_cols)

    for sc in stub_configs:
        q      = sc["question"]
        q_col  = _resolve(q)            # actual df column name
        q_meta = _get_meta(q)

        if q_meta is None:
            continue

        atype        = q_meta["answer_type"]
        choices_i18n = q_meta.get("choices_i18n", {})
        req_stats    = sc.get("stats", ["base", "percent"])

        # Question label: datatable config > question_i18n["en"] > label > q
        q_label = (
            sc.get("label")
            or q_meta.get("question_i18n", {}).get("en")
            or q_meta.get("question_i18n", {}).get("vi")
            or q_meta.get("label")
            or q
        )

        # ── Matrix: expand into one block per sub-question (row) ─────────────
        if atype in MATRIX_TYPES:
            sub_atype = "MA" if atype == "Matrix_MA" else "SA"
            col_choices = choices_i18n.get("columns", {}) if isinstance(choices_i18n, dict) else {}
            for sub_key, sub_meta in q_meta.get("sub_questions", {}).items():
                sub_col   = sub_meta.get("label", sub_key)   # e.g. "Q9_1_r1"
                row_label = sub_meta.get("row_label", sub_col)
                if sub_col not in df.columns:
                    continue
                sub_bases: dict[int, int] = {
                    i: _base_count(df[bc.mask], sub_col)
                    for i, bc in enumerate(banner_cols)
                }
                sub_rows: list[StubRow] = []
                for stat in DEFAULT_STATS_ORDER:
                    if stat not in req_stats:
                        continue
                    if stat == "base":
                        sub_rows.append(StubRow(
                            label="Base", row_type="base",
                            counts=dict(sub_bases),
                            values={i: float(v) for i, v in sub_bases.items()},
                        ))
                    elif stat == "percent" and col_choices:
                        for code in sorted(col_choices.keys(), key=lambda x: int(x)):
                            cnts: dict[int, int]   = {}
                            pcts: dict[int, float] = {}
                            for i, bc in enumerate(banner_cols):
                                sub  = df[bc.mask]
                                base = sub_bases[i]
                                cnt  = _code_count_mc(sub, sub_col, code) if sub_atype == "MA" else _code_count_sc(sub, sub_col, code)
                                cnts[i] = cnt
                                pcts[i] = cnt / base if base else 0.0
                            sub_rows.append(StubRow(
                                label=_choice_label(col_choices, code), row_type="percent",
                                counts=cnts, values=pcts,
                            ))
                if sub_rows:
                    blocks.append(StubBlock(
                        question_code=f"{q.upper()}_R{sub_meta.get('row_index','')}",
                        question_label=f"{q_label} — {row_label}",
                        answer_type=atype,
                        rows=sub_rows,
                    ))
            continue

        if q_col not in df.columns:
            continue

        # base counts keyed by column index
        bases: dict[int, int] = {
            i: _base_count(df[bc.mask], q_col)
            for i, bc in enumerate(banner_cols)
        }

        rows: list[StubRow] = []

        for stat in DEFAULT_STATS_ORDER:
            if stat not in req_stats:
                continue

            if stat == "base":
                rows.append(StubRow(
                    label="Base", row_type="base",
                    counts=dict(bases),
                    values={i: float(v) for i, v in bases.items()},
                ))

            elif stat == "percent":
                if atype not in CODEABLE_TYPES or not choices_i18n:
                    continue
                for code in sorted(choices_i18n.keys(), key=lambda x: int(x)):
                    cnts: dict[int, int]   = {}
                    pcts: dict[int, float] = {}
                    for i, bc in enumerate(banner_cols):
                        sub  = df[bc.mask]
                        base = bases[i]
                        if atype == "MA":
                            cnt = _code_count_mc(sub, q_col, code)
                        else:
                            cnt = _code_count_sc(sub, q_col, code)
                        cnts[i] = cnt
                        pcts[i] = cnt / base if base else 0.0

                    sig = _compute_sig_marks(q_col, code, atype, df, banner_cols, sig_config)
                    rows.append(StubRow(
                        label=_choice_label(choices_i18n, code), row_type="percent",
                        counts=cnts, values=pcts, sig_marks=sig,
                    ))

            elif stat in ("t2b", "b2b"):
                if atype not in CODEABLE_TYPES or not choices_i18n:
                    continue
                sorted_codes = sorted(choices_i18n.keys(), key=lambda x: int(x))
                target = sorted_codes[-2:] if stat == "t2b" else sorted_codes[:2]
                cnts: dict[int, int]   = {}
                pcts: dict[int, float] = {}
                for i, bc in enumerate(banner_cols):
                    sub  = df[bc.mask]
                    base = bases[i]
                    if atype == "MA":
                        cnt = sum(_code_count_mc(sub, q_col, c) for c in target)
                    else:
                        cnt = int(sub[q_col].isin([int(c) for c in target]).sum())
                    cnts[i] = cnt
                    pcts[i] = cnt / base if base else 0.0
                rows.append(StubRow(
                    label=STAT_LABELS[stat], row_type=stat,
                    counts=cnts, values=pcts,
                ))

            elif stat in ("mean", "std", "se"):
                if atype != "SA" or not choices_i18n:
                    continue
                stat_vals: dict[int, float] = {}
                for i, bc in enumerate(banner_cols):
                    numeric = pd.to_numeric(df[bc.mask][q_col], errors="coerce").dropna()
                    if len(numeric) == 0:
                        stat_vals[i] = 0.0
                        continue
                    m = float(numeric.mean())
                    s = float(numeric.std(ddof=1)) if len(numeric) > 1 else 0.0
                    if stat == "mean":
                        stat_vals[i] = round(m, 2)
                    elif stat == "std":
                        stat_vals[i] = round(s, 2)
                    else:
                        stat_vals[i] = round(s / np.sqrt(len(numeric)), 2)
                rows.append(StubRow(
                    label=STAT_LABELS[stat], row_type=stat,
                    counts={i: 0 for i in stat_vals},
                    values=stat_vals,
                ))

        blocks.append(StubBlock(
            question_code=q.upper(),
            question_label=q_label,
            answer_type=atype,
            rows=rows,
        ))

    return blocks
