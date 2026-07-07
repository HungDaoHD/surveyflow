"""Export chart-friendly JSON from table_results.

Transforms the pipeline's table_results (cross-tab format) into a flat,
chart-ready structure consumed by generate_pptx.py and other visualisation tools.
"""
from __future__ import annotations

import json
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

    Returns None if the block has no percent rows (e.g. pure stat blocks).
    """
    rows: list[dict] = block.get("rows", [])
    if not rows:
        return None

    # Collect choices (percent rows only, preserving order)
    choices: list[dict] = []
    choice_codes: list[str] = []
    for row in rows:
        if row.get("row_type") == "percent" and row.get("code") is not None:
            choices.append({"code": str(row["code"]), "label": row["label"]})
            choice_codes.append(str(row["code"]))

    if not choices:
        return None

    # Base row for N counts
    base_row = next((r for r in rows if r.get("row_type") == "base"), None)
    mean_row = next((r for r in rows if r.get("row_type") == "mean"), None)
    nps_row  = next((r for r in rows if r.get("row_type") == "nps"), None)
    # NPS questions (Step 3c: "nps" stat + Promoters/Passives/Detractors
    # choices groups) — the chart shows only these 3 groups instead of every
    # individual code, always in this fixed order (never re-sorted).
    group_rows = [r for r in rows if r.get("row_type") == "group"] if nps_row else []

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
        if group_rows:
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

    def _col_data(idx: int) -> dict:
        data = {"base": _base_at(idx), "percents": _percents_at(idx)}
        mean_v = _stat_at(mean_row, idx)
        if mean_v is not None:
            data["mean"] = mean_v
        nps_v = _stat_at(nps_row, idx)
        if nps_v is not None:
            data["nps"] = nps_v
        return data

    # NPS: replace the individual-code choice list with the 3 fixed groups
    # (Promoters/Passives/Detractors, in the order configured in
    # datatable.json — see CLAUDE.md Step 3c), keyed by "0"/"1"/"2".
    if group_rows:
        choices = [{"code": str(i), "label": row["label"]} for i, row in enumerate(group_rows)]

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

    answer_type = block.get("answer_type", "SA")
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
        "is_nps":     bool(group_rows),
        "choices":    choices,
        "total":      total_data,
        "breakdowns": breakdowns,
    }


# ── Main export function ──────────────────────────────────────────────────────

def export_chart_data(
    table_results: list[dict],
    output_dir: Path,
    survey_name: str = "",
    version: str = "",
    metadata: dict | None = None,
) -> str:
    """Transform table_results → chart_data.json, write to output_dir.

    Args:
        table_results:  context["table_results"] from the pipeline.
        output_dir:     Directory to write chart_data.json (created if missing).
        survey_name:    Embedded in the JSON for reference.
        version:        Pipeline version string (e.g. "v1").
        metadata:       context["metadata"] — used to look up each SA question's
                         Claude-classified `scale_class` ("Ordinal"/"Nominal",
                         see CLAUDE.md Step 3b) by question label.

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

        questions_out: list[dict] = []
        for block in blocks:
            btype = block.get("type", "stub")
            if btype == "stub":
                q = _process_stub(block, total_idx, breakdown_groups, scale_by_label)
                if q:
                    questions_out.append(q)
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
                    if q:
                        questions_out.append(q)
            elif btype == "row_group":
                for sub in block.get("sub_blocks", []):
                    q = _process_stub(sub, total_idx, breakdown_groups, scale_by_label)
                    if q:
                        questions_out.append(q)

        tables_out.append({
            "table_index": tr.get("table_index", 0),
            "title":       tr.get("title", ""),
            "sub_title":   tr.get("sub_title", ""),
            "questions":   questions_out,
        })

    result = {
        "survey":  survey_name,
        "version": version,
        "tables":  tables_out,
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / "chart_data.json"
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(out_path)
