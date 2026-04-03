"""Parse Qme survey rows into a flat list of records (→ rawdata.csv)."""

from __future__ import annotations

from collections import defaultdict
from typing import Any


# ──────────────────────────────────────────────
# Answer serialisers  (row question → str cell)
# ──────────────────────────────────────────────

def _fmt_singlechoice(answer: Any) -> str:
    return str(answer).strip() if answer else ""


def _fmt_multiplechoice(answer: Any) -> str:
    """Returns semicolon-separated answer_names."""
    if not isinstance(answer, list):
        return str(answer).strip() if answer else ""
    return ";".join(a["answer_name"] for a in answer if a.get("answer_name"))


def _fmt_ranking(answer: Any) -> str:
    """Returns ordered semicolon-separated answer_names (rank1 first)."""
    if not isinstance(answer, list):
        return str(answer).strip() if answer else ""
    return ";".join(a["answer_name"] for a in answer if a.get("answer_name"))


def _fmt_matrix(answer: Any) -> str:
    """Returns  row1:col1|row2:col2  string."""
    if not isinstance(answer, list):
        return str(answer).strip() if answer else ""
    parts = []
    for item in answer:
        row = item.get("vertical_answer", "")
        col = item.get("horizontal_answer", "")
        parts.append(f"{row}:{col}")
    return "|".join(parts)


def _fmt_multiplenumber(answer: Any) -> str:
    """Returns  answer1:num1|answer2:num2  string."""
    if not isinstance(answer, list):
        return str(answer).strip() if answer else ""
    parts = []
    for item in answer:
        name = item.get("answer_name", "")
        num  = item.get("number", "")
        parts.append(f"{name}:{num}")
    return "|".join(parts)


def _fmt_photo(q: dict) -> str:
    """Returns semicolon-separated image URLs."""
    images = q.get("images", [])
    return ";".join(images) if images else ""


_SCALAR_TYPES = {
    "user-name", "user-phone", "freetext", "singlechoice",
    "date", "area", "singlenumber", "reward",
}

_ROW_TYPE_TO_FMT = {
    "singlechoice":  _fmt_singlechoice,
    "multiplechoice": _fmt_multiplechoice,
    "ranking":       _fmt_ranking,
    "matrix":        _fmt_matrix,
    "multiplenumber": _fmt_multiplenumber,
}


def _format_answer(rq: dict) -> str:
    """Serialise a single row-question dict to a CSV-safe string."""
    rtype  = rq.get("type", "")
    answer = rq.get("answer", "")

    if rtype == "photo":
        return _fmt_photo(rq)

    if rtype in _SCALAR_TYPES:
        return _fmt_singlechoice(answer)

    fmt = _ROW_TYPE_TO_FMT.get(rtype)
    if fmt:
        return fmt(answer)

    # Fallback: stringify whatever we have
    return str(answer).strip() if answer else ""


# ──────────────────────────────────────────────
# Question-map  (english_question → [position])
# ──────────────────────────────────────────────

def build_question_map(definition_questions: list[dict]) -> dict[str, list[int]]:
    """Build lookup: english_question → ordered list of positions.

    Most questions have a unique english text.  When duplicates exist
    (same text reused for different questions), we keep them in definition
    order so we can match greedily.
    """
    mapping: dict[str, list[int]] = defaultdict(list)
    for q in definition_questions:
        eng = (q.get("english_question") or "").strip()
        if eng:
            mapping[eng].append(q["position"])
    return dict(mapping)


# ──────────────────────────────────────────────
# Row parser
# ──────────────────────────────────────────────

_ROW_META_KEYS = ("task_id", "date_time", "Key_in_date",
                  "lastmodified_date", "profile_status")


def _parse_single_row(row: dict, question_map: dict[str, list[int]]) -> dict:
    """Convert one raw Qme row into a flat {column: value} record."""

    record: dict[str, Any] = {
        "task_id":           row.get("task_id", ""),
        "date_time":         row.get("date_time", ""),
        "Key_in_date":       row.get("Key_in_date", ""),
        "lastmodified_date": row.get("lastmodified_date", ""),
        "profile_status":    row.get("profile_status", ""),
    }

    # Track how many times we've matched each english_question (for duplicates)
    used: dict[str, int] = defaultdict(int)

    for rq in row.get("questions", []):
        eng_q     = (rq.get("question") or "").strip()
        positions = question_map.get(eng_q)
        if not positions:
            continue

        idx = used[eng_q]
        if idx >= len(positions):
            continue  # more answers than definition slots — skip extras

        pos             = positions[idx]
        used[eng_q]    += 1
        label           = f"q{pos}"
        record[label]   = _format_answer(rq)

    return record


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────

def parse_rows(
    rows_pages: list[dict],
    definition: dict,
) -> list[dict]:
    """Parse all fetched row pages into a list of flat records.

    Parameters
    ----------
    rows_pages : list[dict]
        One or more raw ``get_survey_rows`` responses (each page is a
        separate dict).  The ``rows`` list from every page is combined.
    definition : dict
        Raw ``get_survey_definition`` response — used to build the
        question-position map.

    Returns
    -------
    list[dict]
        One flat dict per respondent.  Keys: task_id, date_time,
        Key_in_date, lastmodified_date, profile_status, q1 … qN.
    """
    def_questions = definition.get("questions", [])
    question_map  = build_question_map(def_questions)

    # Collect all instruction positions so we can skip them in the output
    instruction_positions = {
        q["position"]
        for q in def_questions
        if q.get("type") in {31}
    }

    records = []
    for page in rows_pages:
        for row in page.get("rows", []):
            records.append(_parse_single_row(row, question_map))

    return records


def records_to_dataframe(records: list[dict], definition: dict):
    """Convert flat records to a pandas DataFrame with ordered columns.

    Column order: task_id, date_time, Key_in_date, lastmodified_date,
    profile_status, q1, q2, …, qN  (instruction columns excluded).

    Parameters
    ----------
    records : list[dict]
    definition : dict

    Returns
    -------
    pandas.DataFrame
    """
    import pandas as pd

    def_questions = definition.get("questions", [])

    # Build ordered column list (meta + question labels, skip instructions)
    meta_cols = ["task_id", "date_time", "Key_in_date",
                 "lastmodified_date", "profile_status"]
    q_cols = [
        f"q{q['position']}"
        for q in def_questions
        if q.get("type") not in {31}
    ]

    all_cols = meta_cols + q_cols

    df = pd.DataFrame(records)

    # Ensure all expected columns exist (fill missing with "")
    for col in all_cols:
        if col not in df.columns:
            df[col] = ""

    return df[all_cols].fillna("")
