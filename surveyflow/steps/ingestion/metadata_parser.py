"""Parse Qme survey definition into metadata.json structure."""

from __future__ import annotations

import re

# answer_type codes (stored in metadata.json)
_TYPE_LABEL: dict[int, str] = {
    1:    "FT",          # freetext
    2:    "SA",          # singlechoice
    3:    "MA",          # multiplechoice
    4:    "Matrix_SA",   # matrix — single answer per row
    5:    "Matrix_MA",   # matrix — multiple answers per row
    6:    "ranking",
    40:   "record",
    1101: "NUM",         # singlenumber (numeric input)
    1106: "user-name",
    1107: "user-phone",
    1109: "area",
}

_INPUT_TYPE_LABEL: dict[int, str] = {
    3:   "NUM",          # singlenumber input
    68:  "reward",
    100: "multiplenumber",
}

# answer_type codes for sub-questions of each matrix type
_MATRIX_SUB_TYPE: dict[str, str] = {
    "Matrix_SA":  "SA",
    "Matrix_MA":  "MA",
    "Matrix_NUM": "NUM",
}

# answer_types that support numeric coding
CODEABLE_TYPES = {"SA", "MA", "ranking"}

# answer_types excluded from rawdata.csv
EXCLUDED_ANSWER_TYPES = {"audio", "record", "reward"}

# Regex to identify "other-specify" choices — uses word boundaries / lookahead
# to avoid false positives (e.g. "khách" ≠ "khác", "otherwise" ≠ "other")
_OTHER_RE = re.compile(
    r"khác(?!\w)"       # Vietnamese "khác" not followed by a word char  (avoids "khách")
    r"|\bother\b"       # English "other" as whole word
    r"|\bspecify\b"     # "specify"
    r"|\(\s*\)"         # empty/blank parentheses: "( )"
    r"|ghi\s+r[oõ]",   # "ghi rõ" / "ghi ro"
    re.IGNORECASE | re.UNICODE,
)


def _detect_other_codes(choices_i18n: dict) -> list[str]:
    """Return choice codes that are "other-specify" inputs.

    Strategy (in priority order):
    1. ``is_other: true`` field on the choice dict  — explicit, from MCP
    2. Regex pattern match on label text            — fallback for older data
    """
    explicit: list[str] = []
    regex_fallback: list[str] = []

    for code, i18n in choices_i18n.items():
        if not isinstance(i18n, dict):
            continue
        if i18n.get("is_other"):
            explicit.append(str(code))
        else:
            texts = [v for v in i18n.values() if isinstance(v, str)]
            if any(_OTHER_RE.search(t) for t in texts):
                regex_fallback.append(str(code))

    # Use explicit if any found; otherwise fall back to regex
    return explicit if explicit else regex_fallback


def _resolve_answer_type(q_type: int, input_type: int) -> str:
    if q_type == 1 and input_type in _INPUT_TYPE_LABEL:
        return _INPUT_TYPE_LABEL[input_type]
    return _TYPE_LABEL.get(q_type, f"unknown_{q_type}")


def parse_metadata(definition: dict) -> dict:
    """Convert get_survey_definition response → metadata dict.

    Structure:
    {
      "survey_id": ...,
      "questions": {
        "795699": {                          ← key = str(question_id)
          "col":          "q6",             ← rawdata column name
          "position":     6,
          "question_id":  795699,
          "label":        "S1_1",           ← short question label from MCP
          "question_i18n": {"vi": "...", "en": "...", "ja": "..."},
          "answer_type":  "SA",
          "mandatory":    true,
          "status":       1,
          "choices_i18n": {"1": {"vi": "...", "en": "..."}, ...}
        },
        ...
      }
    }
    """
    from surveyflow.steps.ingestion.data_parser import EXCLUDED_ANSWER_TYPES as _EX

    survey = definition["survey"]

    questions: dict[str, dict] = {}
    for q in definition.get("questions", []):
        q_type   = q["type"]
        inp_type = q.get("input_type", 0)
        atype    = _resolve_answer_type(q_type, inp_type)

        if atype in _EX:
            continue

        raw_choices  = q.get("choices", {})
        choices_i18n = q.get("choices_i18n", {})
        pos          = q["position"]
        qid          = q["question_id"]
        col          = f"q{pos}"
        key          = str(qid)

        # ── matrix: build sub_questions (one per row) ────────────────────────
        if atype in _MATRIX_SUB_TYPE and isinstance(raw_choices, dict) and "rows" in raw_choices:
            rows_map    = {str(k): str(v) for k, v in raw_choices["rows"].items()}
            cols_i18n   = {str(k): v for k, v in choices_i18n.get("columns", {}).items()} if isinstance(choices_i18n, dict) else {}
            sub_atype   = _MATRIX_SUB_TYPE[atype]

            # detect "other" column codes once (shared across all rows)
            matrix_other_codes = _detect_other_codes(cols_i18n) if cols_i18n else []

            sub_questions: dict[str, dict] = {}
            for row_code, row_label in rows_map.items():
                sub_key = f"{col}_r{row_code}"
                sub_entry: dict = {
                    "parent":       key,
                    "row_index":    row_code,
                    "answer_type":  sub_atype,
                    "label":        row_label,
                    "choices_i18n": cols_i18n,
                }
                if matrix_other_codes:
                    sub_entry["other_choice_codes"] = matrix_other_codes
                sub_questions[sub_key] = sub_entry
        else:
            matrix_other_codes = []
            sub_questions = {}

        # For regular questions: detect from flat choices_i18n
        # For matrix questions:  reuse matrix_other_codes (detected from cols_i18n above)
        if atype in _MATRIX_SUB_TYPE:
            other_codes = matrix_other_codes   # already computed above
        elif isinstance(choices_i18n, dict) and atype in ("SA", "MA", "ranking"):
            other_codes = _detect_other_codes(choices_i18n)
        else:
            other_codes = []

        q_entry: dict = {
            "position":     pos,
            "question_id":  qid,
            "label":        q.get("label", ""),
            "question_i18n": q.get("question_i18n", {}),
            "answer_type":  atype,
            "mandatory":    q.get("mandatory", False),
            "status":       q.get("status", 1),
            "choices_i18n": choices_i18n,
        }
        if other_codes:
            q_entry["other_choice_codes"] = other_codes
        if sub_questions:
            q_entry["children"]      = list(sub_questions.keys())
            q_entry["sub_questions"] = sub_questions

        questions[key] = q_entry

    return {
        "survey_id":     survey["survey_id"],
        "title":         survey.get("title", ""),
        "english_title": survey.get("english_title", ""),
        "status":        survey.get("status", ""),
        "start_date":    survey.get("start_date", ""),
        "end_date":      survey.get("end_date", ""),
        "questions":     questions,
    }
