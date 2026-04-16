"""Parse Qme survey rows into flat records, then encode answers to numeric codes."""

from __future__ import annotations

import logging
import re
from collections import defaultdict
from typing import Any

logger = logging.getLogger(__name__)


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


def _fmt_location(loc: Any) -> str:
    """[lat, lon] → 'lat,lon'"""
    if isinstance(loc, (list, tuple)) and len(loc) == 2:
        return f"{loc[0]},{loc[1]}"
    return str(loc).strip() if loc else ""


_SCALAR_ROW_TYPES = {
    "user-name", "user-phone", "freetext", "singlechoice",
    "date", "area", "singlenumber", "reward",
}

# Integer type codes that indicate a matrix question in format=code row data
_MATRIX_ROW_TYPES = {4, 5}

_ROW_TYPE_FORMATTERS = {
    "multiplechoice":  _fmt_multiplechoice,
    "ranking":         _fmt_ranking,
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
# Question map
# ─────────────────────────────────────────────

def build_question_id_map(definition_questions: list[dict]) -> dict[int, int]:
    """question_id → position."""
    return {q["question_id"]: q["position"] for q in definition_questions if "question_id" in q}


def _get_other_col_code(
    raw_ans: Any,
    other_codes: list[str],
) -> str | None:
    """Determine the 'other' choice code to use for the other_text column name.

    For SA (scalar answer): the answer value IS the selected code.
    For MA (list answer):   find which selected code is in ``other_codes``.
    Returns the code string (e.g. ``"9"``), or ``None`` when undecidable.
    """
    if not other_codes:
        return None
    if isinstance(raw_ans, (int, float)):
        code = str(int(raw_ans))
        return code if code in other_codes else other_codes[0]
    if isinstance(raw_ans, list):
        # format=code MA: items are dicts with "answer_name" = code string
        for item in raw_ans:
            code = str(item.get("answer_name") or item.get("answer_id") or "").strip()
            if code in other_codes:
                return code
        return other_codes[0]   # fallback: first detected other code
    return other_codes[0] if other_codes else None


# ─────────────────────────────────────────────
# Step 1 – parse rows → raw-text records
# ─────────────────────────────────────────────

_ROW_META = (
    "task_id", "date_time", "Key_in_date", "lastmodified_date",
    "profile_status", "store_id", "check_in", "check_out",
    "user_location", "distance",
)


def _parse_single_row(
    row: dict,
    question_id_map: dict[int, int],
    other_codes_map: dict[int, list[str]] | None = None,
) -> dict:
    record: dict[str, Any] = {
        "task_id":           row.get("task_id", ""),
        "date_time":         row.get("date_time", ""),
        "Key_in_date":       row.get("Key_in_date", ""),
        "lastmodified_date": row.get("lastmodified_date", ""),
        "profile_status":    row.get("profile_status", ""),
        "store_id":          row.get("store-id", ""),
        "check_in":          row.get("check_in", ""),
        "check_out":         row.get("check_out", ""),
        "user_location":     _fmt_location(row.get("user_location")),
        "distance":          row.get("distance", ""),
    }
    used_pos: set[int] = set()
    for rq in row.get("questions", []):
        # ── resolve position by question_id ───────────────────────────────────
        qid = rq.get("question_id")
        pos: int | None = question_id_map.get(qid) if qid else None

        if pos is None:
            if qid:
                logger.warning("question_id=%s not found in definition — skipped", qid)
            continue

        if pos in used_pos:
            continue
        used_pos.add(pos)

        # ── matrix: expand to q{pos}_r{row} sub-columns ──────────────────────
        # format=code returns {"row": int, "col": int}
        rtype_int = rq.get("type")
        if rtype_int in _MATRIX_ROW_TYPES:
            answer_raw = rq.get("answer")
            row_cols: dict[str, list[str]] = defaultdict(list)
            if isinstance(answer_raw, list):
                for item in answer_raw:
                    r = str(item.get("row") or item.get("vertical_answer", "")).strip()
                    c = str(item.get("col") or item.get("horizontal_answer", "")).strip()
                    if r and c:
                        row_cols[r].append(c)
                for row_code, col_codes in row_cols.items():
                    record[f"q{pos}_r{row_code}"] = ";".join(col_codes)

            # ── matrix other_text ──────────────────────────────────────────────
            other = (rq.get("other_text") or "").strip()
            if other and row_cols:
                q_other_codes = (other_codes_map or {}).get(qid, [])
                # Find rows that selected an "other" column code
                other_rows = [
                    (rc, next(c for c in cols if c in q_other_codes))
                    for rc, cols in row_cols.items()
                    if any(c in q_other_codes for c in cols)
                ]
                if other_rows:
                    row_code, other_code = other_rows[0]
                    record[f"q{pos}_r{row_code}_o{other_code}"] = other
                else:
                    record[f"q{pos}_other"] = other   # fallback: can't determine row
            continue

        # ── all other types ───────────────────────────────────────────────────
        answer = _raw_answer(rq)
        record[f"q{pos}"] = answer

        # ── other_text ────────────────────────────────────────────────────────
        other = (rq.get("other_text") or "").strip()
        if other:
            q_other_codes = (other_codes_map or {}).get(qid, [])
            code = _get_other_col_code(rq.get("answer"), q_other_codes)
            col_name = f"q{pos}_o{code}" if code else f"q{pos}_other"
            record[col_name] = other

    return record


def parse_rows(
    rows_pages: list[dict],
    definition: dict,
    profile_status: list[str] | None = None,
    other_codes_map: dict[int, list[str]] | None = None,
) -> list[dict]:
    """Return list of raw records (one per respondent).

    Parameters
    ----------
    profile_status
        Whitelist of statuses to keep. Defaults to ``["approved"]``.
        Pass an empty list ``[]`` to include all statuses.
    other_codes_map
        Mapping ``{question_id: [other_choice_code, ...]}``.
        Used to name other_text columns as ``q{pos}_o{code}``.
        Built from metadata by the caller (IngestionStep).
    """
    if profile_status is None:
        profile_status = ["approved"]

    allowed         = {s.lower() for s in profile_status} if profile_status else None
    def_questions   = definition.get("questions", [])
    question_id_map = build_question_id_map(def_questions)

    records = []
    for page in rows_pages:
        for row in page.get("rows", []):
            if allowed and row.get("profile_status", "").lower() not in allowed:
                continue
            records.append(_parse_single_row(row, question_id_map, other_codes_map))
    return records


# ─────────────────────────────────────────────
# Step 2 – encode answers → numeric codes
# ─────────────────────────────────────────────

from surveyflow.steps.ingestion.metadata_parser import CODEABLE_TYPES


def _parse_number(value: Any) -> int | float | str:
    """Parse formatted number strings: '1,200,000' → 1200000, '32' → 32."""
    if value is None or value == "":
        return ""
    s = str(value).replace(",", "").replace(" ", "").strip()
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        return str(value).strip()


# ── Old-API helpers (used when definition has no `choices`) ───────────────────

def build_encoding_map(
    raw_records: list[dict],
    metadata: dict,
) -> dict[str, dict[str, int]]:
    """Deprecated — kept for backward compatibility; always returns empty dict."""
    return {}


def enrich_metadata_values(
    metadata: dict,
    encoding_map: dict[str, dict[str, int]],
) -> None:
    """Deprecated — kept for backward compatibility; no-op."""
    pass


# ── encode_records ────────────────────────────────────────────────────────────

def _col_to_meta_map(metadata: dict) -> dict[str, dict]:
    """Build {q{position}: metadata_entry} lookup from metadata."""
    return {
        f"q{meta['position']}": meta
        for meta in metadata["questions"].values()
        if "position" in meta
    }


def encode_records(
    raw_records: list[dict],
    metadata: dict,
    encoding_map: dict[str, dict[str, int]] | None = None,   # kept for compat, ignored
) -> list[dict]:
    """Convert answers to numeric codes.

    SA  (singlechoice)   → int            (API already returns integer code)
    MA  (multiplechoice) → "1;3;5"        (API already returns code string)
    ranking              → "2;1;3"        (same)
    NUM                  → int / float    (parse numeric string)
    """
    col_to_meta = _col_to_meta_map(metadata)
    encoded = []

    for record in raw_records:
        new_rec: dict[str, Any] = {}
        for col, value in record.items():
            q = col_to_meta.get(col)
            if q is None or col in _ROW_META or not value:
                new_rec[col] = value
                continue

            atype = q["answer_type"]

            if atype == "SA":
                try:
                    new_rec[col] = int(value)
                except (ValueError, TypeError):
                    new_rec[col] = value
            elif atype == "NUM":
                new_rec[col] = _parse_number(value)
            else:
                new_rec[col] = value

        encoded.append(new_rec)

    return encoded


# ─────────────────────────────────────────────
# Step 3 – records → DataFrame
# ─────────────────────────────────────────────

# answer_types excluded entirely from rawdata.csv
EXCLUDED_ANSWER_TYPES = {"audio", "record", "reward"}


def records_to_dataframe(records: list[dict], definition: dict, metadata: dict):
    """Ordered DataFrame: meta cols + q1…qN (+ _other and matrix sub-cols).

    Question columns are renamed from internal ``q{pos}`` keys to the
    question's ``label`` field.  Duplicate labels raise ``ValueError``.
    """
    import pandas as pd

    col_to_meta = _col_to_meta_map(metadata)

    meta_cols = [
        "task_id", "date_time", "Key_in_date", "lastmodified_date",
        "profile_status", "store_id", "check_in", "check_out",
        "user_location", "distance",
    ]

    _MATRIX_TYPES = {"Matrix_SA", "Matrix_MA", "Matrix_NUM"}

    # Matches:  q7_o9        (regular other)
    #           q7_other     (regular other fallback)
    #           q29_r1_o9    (matrix sub other)
    #           q29_r1_other (matrix sub other fallback)
    _OTHER_COL_RE = re.compile(
        r"^q(\d+)"              # group 1: pos
        r"(?:_r(\d+))?"         # group 2: row (matrix only, optional)
        r"_(o(\d+)|other)$"     # group 3: full suffix, group 4: code (if _o{N})
    )

    # Collect other-text cols, grouped by their "parent" internal col name
    # {parent_col: [other_col, ...]}
    # e.g. "q7"     → ["q7_o9"]
    #      "q29_r1" → ["q29_r1_o9"]
    other_by_parent: dict[str, list[str]] = {}
    for rec in records:
        for k in rec:
            m = _OTHER_COL_RE.match(k)
            if m:
                pos_str = m.group(1)
                row_str = m.group(2)
                parent  = f"q{pos_str}_r{row_str}" if row_str else f"q{pos_str}"
                bucket  = other_by_parent.setdefault(parent, [])
                if k not in bucket:
                    bucket.append(k)

    def _append_other_cols(parent_col: str, parent_label: str) -> None:
        """Append other-text cols that belong to parent_col, with rename entries."""
        for other_col in sorted(other_by_parent.get(parent_col, [])):
            q_cols.append(other_col)
            m = _OTHER_COL_RE.match(other_col)
            code = m.group(4) if m else None
            rename[other_col] = f"{parent_label}_o{code}" if code else f"{parent_label}_other"

    q_cols: list[str] = []          # internal names  (q{pos} / q{pos}_r{row})
    rename: dict[str, str] = {}     # internal → label

    for q in definition.get("questions", []):
        col = f"q{q['position']}"
        meta = col_to_meta.get(col)
        if meta is None:
            continue
        atype  = meta.get("answer_type", "")
        plabel = meta.get("label") or col   # parent label, fallback to q{pos}

        if atype in _MATRIX_TYPES:
            sub_qs = meta.get("sub_questions", {})
            for sub_key in meta.get("children", []):
                q_cols.append(sub_key)
                row_idx   = sub_qs.get(sub_key, {}).get("row_index", sub_key)
                sub_label = f"{plabel}_r{row_idx}"
                rename[sub_key] = sub_label
                _append_other_cols(sub_key, sub_label)
        else:
            q_cols.append(col)
            rename[col] = plabel
            _append_other_cols(col, plabel)

    # ── duplicate label check ─────────────────────────────────────────────────
    label_counts: dict[str, int] = {}
    for lbl in rename.values():
        label_counts[lbl] = label_counts.get(lbl, 0) + 1
    duplicates = sorted(lbl for lbl, cnt in label_counts.items() if cnt > 1)
    if duplicates:
        raise ValueError(
            f"Duplicate question labels detected — rawdata column names must be unique. "
            f"Duplicates: {duplicates}"
        )

    # ── build DataFrame ───────────────────────────────────────────────────────
    all_internal = meta_cols + q_cols

    df = pd.DataFrame(records)
    for col in all_internal:
        if col not in df.columns:
            df[col] = ""
    df = df[all_internal].fillna("")

    # Rename question columns to labels
    df = df.rename(columns=rename)
    all_cols = meta_cols + [rename.get(c, c) for c in q_cols]
    return df[all_cols]
