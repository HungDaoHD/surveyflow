"""Export chart-friendly JSON from table_results.

Transforms the pipeline's table_results (cross-tab format) into a flat,
chart-ready structure consumed by generate_pptx.py and other visualisation tools.
"""
from __future__ import annotations

import json
from pathlib import Path


# ── Chart type heuristics ─────────────────────────────────────────────────────

def _detect_chart_type(answer_type: str, n_choices: int) -> str:
    """Infer chart type from answer_type and number of distinct choices.

    MA / Matrix_MA  → horizontal bar (choices = reasons, may be many)
    SA with ≤ 5     → donut + stacked bar (Likert / scale)
    SA with > 5     → vertical grouped bar (income ranges, age brackets, etc.)
    """
    if answer_type in ("MA", "Matrix_MA"):
        return "bar_horizontal"
    if n_choices <= 5:
        return "donut_stacked"
    return "bar_vertical"


# ── Scale-question detection (Vietnamese + English) ───────────────────────────

# Low/negative anchors expected in the FIRST choice of an ordered rating scale.
# Substring match (anchors are multi-word, low false-positive risk).
_LIKERT_LOW = (
    # Vietnamese
    "hoàn toàn không", "rất không", "không hề", "cực kỳ không",
    "không bao giờ", "hoàn toàn phản đối", "không đồng ý",
    "không liên quan", "không hài lòng", "không quan trọng", "rất tệ",
    # English
    "not at all", "strongly disagree", "very dissatisfied", "extremely dis",
    "never", "completely disagree", "not at all likely",
    "not at all relevant", "very poor", "very unlikely", "not important",
)

# High/positive anchors expected in the LAST choice. Matched at a word boundary
# (start of string or preceded by a space) so "very" doesn't match "every".
_LIKERT_HIGH = (
    # Vietnamese
    "rất ", "hoàn toàn ", "cực kỳ ", "chắc chắn", "luôn luôn",
    "tuyệt vời", "xuất sắc", "vô cùng",
    # English
    "very ", "strongly agree", "extremely ", "completely ", "always",
    "excellent", "definitely", "highly ", "extremely likely",
)


def _detect_scale(labels: list) -> dict:
    """Detect whether an ordered set of choice labels is a rating scale.

    Returns ``{"is_scale", "scale_type", "points"}`` where scale_type is
    ``"numeric"`` (contiguous integers, e.g. 1–5, 0–10), ``"likert"`` (polar
    text anchors at both ends, VN or EN), or ``None``.

    QMe does not tag scale questions, so this is heuristic — works on the choice
    labels alone and is language-agnostic across Vietnamese and English.
    """
    norm = [str(l).strip() for l in labels if str(l).strip()]
    n = len(norm)
    if n < 3:
        return {"is_scale": False, "scale_type": None, "points": n}

    # 1) Numeric scale — labels are contiguous integers (1–5, 0–10, 1–10, …)
    try:
        nums = [int(x) for x in norm]
        if nums == list(range(nums[0], nums[0] + n)):
            return {"is_scale": True, "scale_type": "numeric", "points": n}
    except ValueError:
        pass

    # 2) Likert text scale — polar anchors at both ends (handles up to 11-pt NPS)
    if 3 <= n <= 11:
        first = norm[0].lower()
        last_padded = " " + norm[-1].lower()
        low = any(a in first for a in _LIKERT_LOW)
        high = any((" " + a) in last_padded for a in _LIKERT_HIGH)
        if low and high:
            return {"is_scale": True, "scale_type": "likert", "points": n}

    return {"is_scale": False, "scale_type": None, "points": n}


# ── Block processor ───────────────────────────────────────────────────────────

def _process_stub(block: dict, total_idx: int | None, breakdown_groups: dict) -> dict | None:
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

    def _base_at(idx: int) -> int:
        if base_row is None:
            return 0
        c = base_row.get("counts", [])
        return int(c[idx]) if idx < len(c) else 0

    def _percents_at(idx: int) -> dict[str, float]:
        out: dict[str, float] = {}
        for row in rows:
            if row.get("row_type") == "percent" and row.get("code") is not None:
                vals = row.get("values", [])
                v = float(vals[idx]) if idx < len(vals) else 0.0
                out[str(row["code"])] = round(v, 4)
        return out

    # Total column
    total_data: dict = {}
    if total_idx is not None:
        total_data = {
            "base":    _base_at(total_idx),
            "percents": _percents_at(total_idx),
        }

    # Breakdown groups (non-total banner columns, grouped by group_label)
    breakdowns: list[dict] = []
    for group_label, cols in breakdown_groups.items():
        columns = []
        for bc in cols:
            idx = bc["_idx"]
            columns.append({
                "label":    bc.get("label", ""),
                "base":     _base_at(idx),
                "percents": _percents_at(idx),
            })
        breakdowns.append({"group_label": group_label, "columns": columns})

    answer_type = block.get("answer_type", "SA")
    scale = _detect_scale([c["label"] for c in choices])
    return {
        "question":   block.get("question", ""),
        "label":      block.get("label", ""),
        "answer_type": answer_type,
        "chart_type": _detect_chart_type(answer_type, len(choices)),
        "scale":      scale,
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
) -> str:
    """Transform table_results → chart_data.json, write to output_dir.

    Args:
        table_results:  context["table_results"] from the pipeline.
        output_dir:     Directory to write chart_data.json (created if missing).
        survey_name:    Embedded in the JSON for reference.
        version:        Pipeline version string (e.g. "v1").

    Returns:
        Absolute path of the written file.
    """
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
                q = _process_stub(block, total_idx, breakdown_groups)
                if q:
                    questions_out.append(q)
            elif btype in ("row_group", "ranking"):
                for sub in block.get("sub_blocks", []):
                    q = _process_stub(sub, total_idx, breakdown_groups)
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
