"""Export chart-friendly JSON from table_results.

Transforms the pipeline's table_results (cross-tab format) into a flat,
chart-ready structure consumed by generate_pptx.py and other visualisation tools.
"""
from __future__ import annotations

import json
import re
from pathlib import Path


# ── Chart type heuristics ─────────────────────────────────────────────────────

def _detect_chart_type(answer_type: str, n_choices: int) -> str:
    """Infer chart type from answer_type.

    MA / Matrix_MA → horizontal bar (choices = reasons, may be many)
    Everything else (SA, Matrix_SA, ...) → donut + stacked bar, regardless of
    choice count — all SA questions now always use the donut/stack layout.
    """
    if answer_type in ("MA", "Matrix_MA"):
        return "bar_horizontal"
    return "donut_stacked"


# ── Scale-question detection (Vietnamese + English) ───────────────────────────

# Low/negative anchors expected in the lowest-code choice of an ordered rating
# scale. Substring match (anchors are multi-word, low false-positive risk).
_LIKERT_LOW = (
    # Vietnamese
    "hoàn toàn không", "rất không", "không hề", "cực kỳ không",
    "không bao giờ", "hoàn toàn phản đối", "không đồng ý",
    "không liên quan", "không hài lòng", "không quan trọng", "rất tệ",
    "chắc chắn không mua", "sẽ không mua", "không chắc",
    # English
    "not at all", "strongly disagree", "very dissatisfied", "extremely dis",
    "never", "completely disagree", "not at all likely",
    "not at all relevant", "very poor", "very unlikely", "not important",
    "definitely will not", "definitely won't", "definitely wont",
)

# High/positive anchors expected in the highest-code choice. Matched at a word
# boundary (start of string or preceded by a space) so "very" doesn't match
# "every".
_LIKERT_HIGH = (
    # Vietnamese
    "rất ", "hoàn toàn ", "cực kỳ ", "chắc chắn", "luôn luôn",
    "tuyệt vời", "xuất sắc", "vô cùng",
    # English
    "very ", "strongly agree", "extremely ", "completely ", "always",
    "excellent", "definitely", "highly ", "extremely likely",
)


def _detect_scale(choices: list) -> dict:
    """Detect whether a set of choices (``{"code", "label"}`` dicts) is a
    rating scale.

    Returns ``{"is_scale", "scale_type", "points"}`` where scale_type is
    ``"likert"`` (contiguous codes with polar anchor text at the low/high end,
    VN or EN) or ``"numeric"`` (contiguous codes, no anchor text), or ``None``.

    Uses the answer **code** (not label text or list position) as the primary
    signal — chart_data's choice order isn't guaranteed to already be sorted
    ascending by scale value, and labels can mix plain numbers with full
    anchor text only at the endpoints (e.g. "3", "4", "2",
    "5 – Definitely will buy", "1 – Definitely won't buy"). QMe does not tag
    scale questions, so this is heuristic and language-agnostic across
    Vietnamese and English.
    """
    n = len(choices)
    if n < 3:
        return {"is_scale": False, "scale_type": None, "points": n}

    try:
        codes = [int(c["code"]) for c in choices]
    except (ValueError, TypeError, KeyError):
        return {"is_scale": False, "scale_type": None, "points": n}

    # Codes must be a contiguous integer range (e.g. 1–5, 0–10, 1–10, …),
    # regardless of what order the choices happen to be listed in.
    if sorted(codes) != list(range(min(codes), min(codes) + n)):
        return {"is_scale": False, "scale_type": None, "points": n}

    if n > 11:
        return {"is_scale": True, "scale_type": "numeric", "points": n}

    by_code = sorted(choices, key=lambda c: int(c["code"]))
    low_label  = str(by_code[0]["label"]).strip().lower()
    high_label = " " + str(by_code[-1]["label"]).strip().lower()
    low  = any(a in low_label for a in _LIKERT_LOW)
    high = any((" " + a) in high_label for a in _LIKERT_HIGH)
    scale_type = "likert" if (low and high) else "numeric"
    return {"is_scale": True, "scale_type": scale_type, "points": n}


# ── Block processor ───────────────────────────────────────────────────────────

def _process_stub(block: dict, total_idx: int | None, breakdown_groups: dict,
                   scale_by_label: dict | None = None) -> dict | None:
    """Convert a single StubBlock dict → chart question dict.

    Returns None if the block has no usable choices (e.g. pure stat blocks).
    """
    rows: list[dict] = block.get("rows", [])
    if not rows:
        return None

    answer_type = block.get("answer_type", "SA")

    # Collect choices from individual percent rows, preserving order.
    choices: list[dict] = []
    for row in rows:
        if row.get("row_type") == "percent" and row.get("code") is not None:
            choices.append({"code": str(row["code"]), "label": row["label"]})

    base_row = next((r for r in rows if r.get("row_type") == "base"), None)
    nps_row  = next((r for r in rows if r.get("row_type") == "nps"), None)
    group_rows = [r for r in rows if r.get("row_type") == "group"]

    # NPS (Step 3c: "nps" stat + Promoters/Passives/Detractors groups) always
    # uses its 3 fixed groups instead of individual codes.
    #
    # Range/bucket groups — NUM using "num_quantile" (see CLAUDE.md Step 4
    # NUM defaults), or any question whose "choices" config hides every
    # individual code and shows only group summaries: when there are NO
    # percent rows of its own but there ARE "group" rows, use those groups
    # as the chart's choice list — otherwise the question has zero choices
    # and would silently vanish from the appendix (no slide at all).
    use_group_as_choices = bool(nps_row) or (not choices and bool(group_rows))
    if use_group_as_choices:
        choices = [{"code": str(i), "label": row["label"]} for i, row in enumerate(group_rows)]

    # multiplenumber (budget-allocation style questions, e.g. "how much did
    # you spend on each category"): "percent" rows are response RATES (% who
    # entered ANY value for that category — don't sum to 100%), not a valid
    # donut/100%-stack source. "mean" rows use a common-denominator formula
    # (see table_generator.py Step 3c) that DOES sum to the respondent-level
    # constant total by design (10, 100, whatever the survey enforces) —
    # normalize them to fractions-of-1 and use THOSE as the chart's percents
    # instead, so multiplenumber renders as donut+stack exactly like an
    # SA-Ordinal question (per-category share of the average allocation).
    mean_rows_by_code: dict[str, dict] = {}
    if answer_type == "multiplenumber":
        mean_rows_by_code = {
            str(r["code"]): r for r in rows
            if r.get("row_type") == "mean" and r.get("code") is not None
        }
        if mean_rows_by_code:
            label_by_code = {c["code"]: c["label"] for c in choices}
            choices = [
                {"code": code, "label": label_by_code.get(code, r["label"])}
                for code, r in mean_rows_by_code.items()
            ]

    if not choices:
        return None

    def _base_at(idx: int) -> int:
        if base_row is None:
            return 0
        c = base_row.get("counts", [])
        return int(c[idx]) if idx < len(c) else 0

    def _stat_at(row: dict | None, idx: int) -> float | None:
        if row is None:
            return None
        vals = row.get("values", [])
        return round(float(vals[idx]), 2) if idx < len(vals) else None

    def _percents_at(idx: int) -> dict[str, float]:
        out: dict[str, float] = {}
        if mean_rows_by_code:
            raw = {code: (r.get("values", [])[idx] if idx < len(r.get("values", [])) else 0.0)
                   for code, r in mean_rows_by_code.items()}
            total = sum(raw.values())
            for code, v in raw.items():
                out[code] = round(v / total, 4) if total else 0.0
            return out
        if use_group_as_choices:
            for i, row in enumerate(group_rows):
                vals = row.get("values", [])
                out[str(i)] = round(float(vals[idx]), 4) if idx < len(vals) else 0.0
            return out
        for row in rows:
            if row.get("row_type") == "percent" and row.get("code") is not None:
                vals = row.get("values", [])
                v = float(vals[idx]) if idx < len(vals) else 0.0
                out[str(row["code"])] = round(v, 4)
        return out

    # A single scalar "Mean: X" donut-hole overlay (Step 3c, SA/Matrix_SA
    # Ordinal / NUM) only makes sense when there's exactly one mean value for
    # the whole question. multiplenumber has one mean PER CATEGORY (already
    # used as the percents source above) — showing any one of them as if it
    # were "the" question's mean would be misleading, so skip the scalar
    # overlay entirely for multiplenumber.
    mean_row = None if mean_rows_by_code else next(
        (r for r in rows if r.get("row_type") == "mean"), None)

    def _col_data(idx: int) -> dict:
        data = {"base": _base_at(idx), "percents": _percents_at(idx)}
        mean_v = _stat_at(mean_row, idx)
        if mean_v is not None:
            data["mean"] = mean_v
        nps_v = _stat_at(nps_row, idx)
        if nps_v is not None:
            data["nps"] = nps_v
        return data

    # Total column
    total_data: dict = {}
    if total_idx is not None:
        total_data = _col_data(total_idx)

    # Breakdown groups (non-total banner columns, grouped by group_label)
    breakdowns: list[dict] = []
    for group_label, cols in breakdown_groups.items():
        columns = []
        for bc in cols:
            idx = bc["_idx"]
            columns.append({"label": bc.get("label", ""), **_col_data(idx)})
        breakdowns.append({"group_label": group_label, "columns": columns})

    # Match on `question` (short code, e.g. "C4") not `label` — `label` is the
    # datatable stub's display label, which falls back to the FULL question
    # text (question_i18n) when the stub doesn't set a custom one, so it
    # rarely matches metadata's short `label` field used as the lookup key.
    scale_info = (scale_by_label or {}).get(block.get("question", "").upper(), {})
    return {
        "question":   block.get("question", ""),
        "label":      block.get("label", ""),
        "answer_type": answer_type,
        "chart_type": _detect_chart_type(answer_type, len(choices)),
        "scale_class":     scale_info.get("scale_class"),
        "scale_high_code": scale_info.get("scale_high_code"),
        "is_nps":     bool(nps_row),
        # Range/bucket groups (NUM num_quantile / all-hidden "choices") must
        # keep their natural emitted order (e.g. ascending age bands) rather
        # than being re-sorted by Total percent like a Nominal SA question.
        "preserve_order": use_group_as_choices and not nps_row,
        "choices":    choices,
        "total":      total_data,
        "breakdowns": breakdowns,
    }


# ── Slide title override ──────────────────────────────────────────────────────
#
# `datatable.json` stub entries may carry an optional user-authored "title" —
# a short summary to show as the slide's headline instead of the (often long,
# verbatim survey-wording) question label. It's authored on the stub config
# (the file Claude/the user actually edits), not on chart_data.json directly,
# then flattened here into a question-code → title map and attached to each
# exported question by matching codes — see _lookup_title for how a title set
# once on a matrix/ranking parent question covers every row/rank slide it
# expands into.

def _flatten_stub_titles(stub_configs: list[dict]) -> dict[str, str]:
    """question-code (upper) -> title, from every stub entry's own "title"
    plus one level of row_group "items" (each item is itself a stub-like
    config with its own "question"/"title")."""
    out: dict[str, str] = {}
    for sc in stub_configs:
        title = sc.get("title")
        q = str(sc.get("question", "")).strip().upper()
        if title and q:
            out[q] = title
        for item in sc.get("items", []) or []:
            i_title = item.get("title")
            i_q = str(item.get("question", "")).strip().upper()
            if i_title and i_q:
                out[i_q] = i_title
    return out


_ROW_SUFFIX_RE  = re.compile(r"_R\d+$")
_RANK_SUFFIX_RE = re.compile(r"-(RANK\s*\d+|OVERALL)$")


def _lookup_title(title_by_qcode: dict[str, str], question_code: str) -> str | None:
    """Exact match first; else strip a trailing matrix-row suffix ("_R10") or
    ranking sub-block suffix ("-Rank 1"/"-Overall") and retry against the
    parent question's own code — so a title set once on a matrix/ranking stub
    entry applies to every row/rank slide it expands into, not just one."""
    q = str(question_code or "").strip().upper()
    if q in title_by_qcode:
        return title_by_qcode[q]
    base = _RANK_SUFFIX_RE.sub("", _ROW_SUFFIX_RE.sub("", q))
    return title_by_qcode.get(base)


def _flatten_meta_titles(metadata: dict | None, lang: str) -> dict[str, str]:
    """question-code (upper) -> title, from metadata.json's AI-summarized
    `title_i18n` (see CLAUDE.md Step 3a) — the fallback source for "title"
    when the stub config doesn't set one, mirroring exactly how "label"
    already falls back to metadata's `question_i18n` in table_generator.py.
    Matrix sub_questions carry their own `title_i18n` (parent title + row
    label, already combined by Step 3a), so no suffix-stripping is needed
    here — only _lookup_title's ranking-suffix fallback needs it."""
    out: dict[str, str] = {}
    if not metadata:
        return out

    def _add(label: str | None, title_i18n: dict | None) -> None:
        if not label or not title_i18n:
            return
        text = title_i18n.get(lang) or title_i18n.get("en") or title_i18n.get("vi")
        if text:
            out[str(label).strip().upper()] = text

    for meta in metadata.get("questions", {}).values():
        _add(meta.get("label"), meta.get("title_i18n"))
        for sub in (meta.get("sub_questions") or {}).values():
            _add(sub.get("label"), sub.get("title_i18n"))
    return out


# ── Main export function ──────────────────────────────────────────────────────

def export_chart_data(
    table_results: list[dict],
    output_dir: Path,
    survey_name: str = "",
    version: str = "",
    metadata: dict | None = None,
    lang: str = "vi",
) -> str:
    """Transform table_results → chart_data.json, write to output_dir.

    Args:
        table_results:  context["table_results"] from the pipeline.
        output_dir:     Directory to write chart_data.json (created if missing).
        survey_name:    Embedded in the JSON for reference.
        version:        Pipeline version string (e.g. "v1").
        metadata:       context["metadata"] — used to look up each SA question's
                         Claude-classified `scale_class` ("Ordinal"/"Nominal",
                         see CLAUDE.md Step 3b) by question label, and each
                         question's AI-summarized `title_i18n` (Step 3a) as the
                         fallback source for "title" when the stub config
                         (datatable.json) doesn't set one — same
                         override-then-fallback pattern "label" already uses.
        lang:           Display language used for this run's labels (e.g. "vi"/
                         "en") — embedded in the JSON so generate_pptx.py can
                         pick a language-appropriate Dzung_team section tag
                         ("PHỤ LỤC" vs "APPENDIX") without a separate CLI flag,
                         and selects which `title_i18n` language to fall back to.

    Returns:
        Absolute path of the written file.
    """
    scale_by_label: dict[str, dict] = {}
    if metadata:
        def _add(label: str | None, sc: str | None, high_code) -> None:
            if label and sc:
                scale_by_label[label.upper()] = {
                    "scale_class": sc, "scale_high_code": high_code,
                }

        for meta in metadata.get("questions", {}).values():
            _add(meta.get("label"), meta.get("scale_class"), meta.get("scale_high_code"))
            # Matrix questions: sub_questions (one per row, e.g. "A4_r1") carry
            # their own scale_class/scale_high_code, since chart_data blocks
            # for matrix rows are keyed by the sub-question's code, not the
            # parent's.
            for sub in (meta.get("sub_questions") or {}).values():
                _add(sub.get("label"), sub.get("scale_class"), sub.get("scale_high_code"))

    title_by_meta = _flatten_meta_titles(metadata, lang)

    tables_out: list[dict] = []

    for tr in table_results:
        banner_cols: list[dict] = tr.get("banner_cols", [])
        blocks:      list[dict] = tr.get("blocks", [])

        # Identify total column index; group remaining cols by group_label (order-preserving)
        total_idx: int | None = None
        breakdown_groups: dict[str, list[dict]] = {}

        for i, bc in enumerate(banner_cols):
            if bc.get("is_total"):
                total_idx = i
            else:
                g = bc.get("group_label", "")
                bc_entry = {**bc, "_idx": i}
                breakdown_groups.setdefault(g, []).append(bc_entry)

        title_by_qcode = _flatten_stub_titles(tr.get("stub", []))

        questions_out: list[dict] = []

        def _append(q: dict | None) -> None:
            if not q:
                return
            # Stub-authored title (datatable.json) wins if set; otherwise fall
            # back to metadata.json's AI-summarized title_i18n — same
            # override-then-fallback pattern "label" already uses against
            # question_i18n. Rebuilt (not mutated) so "title" lands right
            # after "question" in the JSON, not appended at the end.
            code = q.get("question", "")
            title = _lookup_title(title_by_qcode, code) or _lookup_title(title_by_meta, code)
            q = {"question": code, "title": title,
                 **{k: v for k, v in q.items() if k not in ("question", "title")}}
            questions_out.append(q)

        for block in blocks:
            btype = block.get("type", "stub")
            if btype == "stub":
                q = _process_stub(block, total_idx, breakdown_groups, scale_by_label)
                _append(q)
            elif btype == "ranking":
                # Sub-blocks are per-rank-position ("Rank 1", "Rank 2", …) or a
                # single flat "any rank" block — neither carries the actual
                # ranking question's own code/wording (only the parent
                # RankingBlock does: "question"="Q27_2", "label"=full text),
                # so each slide would otherwise show just "RANK1." / "Rank 1"
                # with no link back to the source question. Combine the
                # parent's question code + label onto every sub-block before
                # rendering, e.g. "Q27_2-Rank 1".
                parent_q     = block.get("question", "")
                parent_label = block.get("label", "")
                mode = block.get("mode", "rank_dist")
                for rank_num, sub in enumerate(block.get("sub_blocks", []), start=1):
                    suffix = f"Rank {rank_num}" if mode == "rank_dist" else "Overall"
                    sub = {
                        **sub,
                        "question": f"{parent_q}-{suffix}" if parent_q else suffix,
                        "label": f"{parent_label} — {suffix}" if parent_label else suffix,
                    }
                    if mode == "any_rank":
                        # any_rank's rows are "% ranked at ANY position", one
                        # independent percent per item (not mutually
                        # exclusive, don't sum to 100%) — same shape as an MA
                        # question, so it must chart/label as MA (bar chart),
                        # not the donut used for the per-rank-position blocks.
                        sub["answer_type"] = "MA"
                    q = _process_stub(sub, total_idx, breakdown_groups, scale_by_label)
                    _append(q)
            elif btype == "row_group":
                for sub in block.get("sub_blocks", []):
                    q = _process_stub(sub, total_idx, breakdown_groups, scale_by_label)
                    _append(q)

        tables_out.append({
            "table_index":     tr.get("table_index", 0),
            "title":           tr.get("title", ""),
            "sub_title":       tr.get("sub_title", ""),
            "appendix_format": tr.get("appendix_format", "general"),
            "questions":       questions_out,
        })

    result = {
        "survey":  survey_name,
        "version": version,
        "lang":    lang,
        "tables":  tables_out,
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / "chart_data.json"
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(out_path)
