"""Parse Qme survey rows into flat records, then encode answers to numeric codes."""

from __future__ import annotations

from collections import defaultdict
from typing import Any


# ─────────────────────────────────────────────
# Step 1 – raw-text serialisers (row → str)
# ─────────────────────────────────────────────

def _fmt_text(answer: Any) -> str:
    return str(answer).strip() if answer else ""


def _fmt_multiplechoice(answer: Any) -> str:
    """Semicolon-separated answer_names."""
    if not isinstance(answer, list):
        return _fmt_text(answer)
    return ";".join(a["answer_name"] for a in answer if a.get("answer_name"))


def _fmt_ranking(answer: Any) -> str:
    """Ordered semicolon-separated answer_names (rank-1 first)."""
    if not isinstance(answer, list):
        return _fmt_text(answer)
    return ";".join(a["answer_name"] for a in answer if a.get("answer_name"))


def _fmt_matrix(answer: Any) -> str:
    """row1:col1|row2:col2"""
    if not isinstance(answer, list):
        return _fmt_text(answer)
    return "|".join(
        f"{i.get('vertical_answer','')}:{i.get('horizontal_answer','')}"
        for i in answer
    )


def _fmt_multiplenumber(answer: Any) -> str:
    """answer1:num1|answer2:num2"""
    if not isinstance(answer, list):
        return _fmt_text(answer)
    return "|".join(
        f"{i.get('answer_name','')}:{i.get('number','')}"
        for i in answer
    )


def _fmt_photo(rq: dict) -> str:
    return ";".join(rq.get("images", []))


_SCALAR_ROW_TYPES = {
    "user-name", "user-phone", "freetext", "singlechoice",
    "date", "area", "singlenumber", "reward",
}

_ROW_TYPE_FORMATTERS = {
    "multiplechoice":  _fmt_multiplechoice,
    "ranking":         _fmt_ranking,
    "matrix":          _fmt_matrix,
    "multiplenumber":  _fmt_multiplenumber,
}


def _raw_answer(rq: dict) -> str:
    rtype = rq.get("type", "")
    if rtype == "photo":
        return _fmt_photo(rq)
    if rtype in _SCALAR_ROW_TYPES:
        return _fmt_text(rq.get("answer", ""))
    fmt = _ROW_TYPE_FORMATTERS.get(rtype)
    if fmt:
        return fmt(rq.get("answer", ""))
    return _fmt_text(rq.get("answer", ""))


# ─────────────────────────────────────────────
# Question map  (english_question → [position])
# ─────────────────────────────────────────────

def build_question_map(definition_questions: list[dict]) -> dict[str, list[int]]:
    mapping: dict[str, list[int]] = defaultdict(list)
    for q in definition_questions:
        eng = (q.get("english_question") or "").strip()
        if eng:
            mapping[eng].append(q["position"])
    return dict(mapping)


# ─────────────────────────────────────────────
# Step 1 – parse rows → raw-text records
# ─────────────────────────────────────────────

_ROW_META = ("task_id", "date_time", "Key_in_date",
             "lastmodified_date", "profile_status")


def _parse_single_row(row: dict, question_map: dict[str, list[int]]) -> dict:
    record: dict[str, Any] = {
        "task_id":           row.get("task_id", ""),
        "date_time":         row.get("date_time", ""),
        "Key_in_date":       row.get("Key_in_date", ""),
        "lastmodified_date": row.get("lastmodified_date", ""),
        "profile_status":    row.get("profile_status", ""),
    }
    used: dict[str, int] = defaultdict(int)
    for rq in row.get("questions", []):
        eng_q     = (rq.get("question") or "").strip()
        positions = question_map.get(eng_q)
        if not positions:
            continue
        idx = used[eng_q]
        if idx >= len(positions):
            continue
        pos           = positions[idx]
        used[eng_q]  += 1
        record[f"q{pos}"] = _raw_answer(rq)
    return record


def parse_rows(
    rows_pages: list[dict],
    definition: dict,
    profile_status: list[str] | None = None,
) -> list[dict]:
    """Return list of raw-text records (one per respondent).

    Parameters
    ----------
    profile_status
        Whitelist of statuses to keep. Defaults to ``["approved"]``.
        Pass an empty list ``[]`` to include all statuses.
    """
    if profile_status is None:
        profile_status = ["approved"]

    allowed = {s.lower() for s in profile_status} if profile_status else None
    question_map = build_question_map(definition.get("questions", []))
    records = []
    for page in rows_pages:
        for row in page.get("rows", []):
            if allowed and row.get("profile_status", "").lower() not in allowed:
                continue
            records.append(_parse_single_row(row, question_map))
    return records


# ─────────────────────────────────────────────
# Step 2 – encode answers → numeric codes
# ─────────────────────────────────────────────

from surveyflow.steps.ingestion.metadata_parser import CODEABLE_TYPES


# ── Old-API helpers (used when definition has no `choices`) ───────────────────

def build_encoding_map(
    raw_records: list[dict],
    metadata: dict,
) -> dict[str, dict[str, int]]:
    """Build {col: {answer_text → 1-based code}} from first-appearance order.

    Only processes questions whose ``values`` dict is still empty (i.e. the
    definition did not supply ``choices``).
    """
    questions = metadata["questions"]
    seen: dict[str, dict[str, int]] = {}

    for record in raw_records:
        for col, value in record.items():
            if not value:
                continue
            q = questions.get(col)
            if q is None or q["answer_type"] not in CODEABLE_TYPES:
                continue
            if q.get("values"):          # already populated from definition.choices
                continue
            if col not in seen:
                seen[col] = {}
            atype = q["answer_type"]
            if atype == "multiplechoice":
                for part in str(value).split(";"):
                    part = part.strip()
                    if part and part not in seen[col]:
                        seen[col][part] = len(seen[col]) + 1
            else:                        # singlechoice / ranking
                val = str(value).strip()
                if val and val not in seen[col]:
                    seen[col][val] = len(seen[col]) + 1

    return seen


def enrich_metadata_values(
    metadata: dict,
    encoding_map: dict[str, dict[str, int]],
) -> None:
    """Fill ``metadata["questions"][col]["values"]`` from encoding_map.

    Result format: ``{code_str: label_text}``  e.g. ``{"1": "Male", "2": "Female"}``
    """
    for col, label_to_code in encoding_map.items():
        q = metadata["questions"].get(col)
        if q is not None:
            q["values"] = {str(code): label for label, code in label_to_code.items()}


# ── encode_records ────────────────────────────────────────────────────────────

def encode_records(
    raw_records: list[dict],
    metadata: dict,
    encoding_map: dict[str, dict[str, int]] | None = None,
) -> list[dict]:
    """Convert answers to numeric codes.

    New API (encoding_map is None):
      singlechoice   → int            (API already returns integer code)
      multiplechoice → "1;3;5"        (API already returns code string)
      ranking        → "2;1;3"        (same)

    Old API (encoding_map provided):
      singlechoice   → int code       (looked up from encoding_map)
      multiplechoice → "1;3;5"        (each text part looked up)
      ranking        → same
    """
    questions = metadata["questions"]
    encoded = []

    for record in raw_records:
        new_rec: dict[str, Any] = {}
        for col, value in record.items():
            q = questions.get(col)
            if q is None or col in _ROW_META or not value:
                new_rec[col] = value
                continue

            atype = q["answer_type"]

            if encoding_map is not None and col in encoding_map:
                # Old-API path: convert text → code
                label_to_code = encoding_map[col]
                if atype == "singlechoice":
                    new_rec[col] = label_to_code.get(str(value).strip(), "")
                else:   # multiplechoice / ranking
                    parts = [v.strip() for v in str(value).split(";") if v.strip()]
                    codes = [str(label_to_code[p]) for p in parts if p in label_to_code]
                    new_rec[col] = ";".join(codes)
            elif atype == "singlechoice":
                # New-API path: cast to int
                try:
                    new_rec[col] = int(value)
                except (ValueError, TypeError):
                    new_rec[col] = value
            else:
                new_rec[col] = value

        encoded.append(new_rec)

    return encoded


# ─────────────────────────────────────────────
# Step 3 – records → DataFrame
# ─────────────────────────────────────────────

# answer_types excluded from rawdata.csv
EXCLUDED_ANSWER_TYPES = {"audio", "user-name", "user-phone", "instruction", "reward"}


def records_to_dataframe(records: list[dict], definition: dict, metadata: dict):
    """Ordered DataFrame: meta cols + q1…qN.
    Only includes q-columns that exist in metadata (single source of truth).
    """
    import pandas as pd

    questions_meta = metadata.get("questions", {})

    meta_cols = ["task_id", "date_time", "Key_in_date",
                 "lastmodified_date", "profile_status"]
    q_cols = [
        f"q{q['position']}"
        for q in definition.get("questions", [])
        if f"q{q['position']}" in questions_meta
    ]
    all_cols = meta_cols + q_cols

    df = pd.DataFrame(records)
    for col in all_cols:
        if col not in df.columns:
            df[col] = ""
    return df[all_cols].fillna("")
