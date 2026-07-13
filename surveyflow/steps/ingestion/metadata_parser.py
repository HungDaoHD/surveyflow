"""Parse Qme survey definition into metadata.json structure."""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)

# answer_type codes (stored in metadata.json)
_TYPE_LABEL: dict[int, str] = {
    1:    "FT",          # freetext
    2:    "SA",          # singlechoice
    3:    "MA",          # multiplechoice
    4:    "Matrix_SA",   # matrix — single answer per row
    5:    "Matrix_MA",   # matrix — multiple answers per row
    6:    "ranking",
    8:    "SA",          # singlechoice (alternate type code)
    9:    "MA",          # grid MA (answer coded same as MA)
    28:   "Matrix_SA",   # matrix SA — numeric/rating per row
    29:   "Matrix_MA",   # matrix MA — multiple cols per row
    31:   "instruction", # section header / description block
    40:   "record",
    1100: "instruction", # quota / screener field
    1101: "NUM",         # singlenumber (numeric input)
    1106: "user-name",
    1107: "user-phone",
    1109: "SA",          # area → treated as SA with synthetic choices
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
EXCLUDED_ANSWER_TYPES = {"audio", "record", "reward", "instruction", "user-name", "user-phone", "photo"}

def _build_lookup_maps(
    questions: list[dict],
) -> tuple[dict[str, str], dict[int, tuple[str, int]], dict[int, tuple[str, str]], set[str]]:
    """Build lookup maps from the raw definition questions list.

    Returns
    -------
    id_to_label : {str(question_id): question_label}
    answer_id_to_code : {answer_id: (question_label, int_choice_code)}
        Flat choices (SA/MA/ranking) + matrix column choices.
    row_answer_id_to_info : {row_answer_id: (question_label, row_code_str)}
        Matrix row choices — used to decode compound IDs like "799010-5582321".
    q_ids_with_answer_ids : set[str]
        Question IDs whose choices carry ``answer_id`` fields.
        Questions NOT in this set (e.g. area/special types) use direct choice
        codes in condition rules — no answer_id translation needed.
    """
    id_to_label: dict[str, str] = {}
    answer_id_to_code: dict[int, tuple[str, int]] = {}
    row_answer_id_to_info: dict[int, tuple[str, str]] = {}
    q_ids_with_answer_ids: set[str] = set()

    for q in questions:
        qid   = str(q.get("question_id", ""))
        label = q.get("label", "")
        if qid:
            id_to_label[qid] = label

        ci = q.get("choices_i18n", {})
        if not isinstance(ci, dict):
            continue

        # Flat choices (SA / MA / ranking)
        for code_str, ch_val in ci.items():
            if isinstance(ch_val, dict) and "answer_id" in ch_val:
                try:
                    answer_id_to_code[int(ch_val["answer_id"])] = (label, int(code_str))
                    q_ids_with_answer_ids.add(qid)
                except (ValueError, TypeError):
                    pass

        # Matrix column choices → used to translate column answer_id → code
        for code_str, ch_val in (ci.get("columns") or {}).items():
            if isinstance(ch_val, dict) and "answer_id" in ch_val:
                try:
                    answer_id_to_code[int(ch_val["answer_id"])] = (label, int(code_str))
                    q_ids_with_answer_ids.add(qid)
                except (ValueError, TypeError):
                    pass

        # Matrix row choices → used to decode compound id "{q_id}-{row_answer_id}"
        for row_code_str, row_val in (ci.get("rows") or {}).items():
            if isinstance(row_val, dict) and "answer_id" in row_val:
                try:
                    row_answer_id_to_info[int(row_val["answer_id"])] = (label, row_code_str)
                except (ValueError, TypeError):
                    pass

    return id_to_label, answer_id_to_code, row_answer_id_to_info, q_ids_with_answer_ids


def _translate_condition_node(
    node: "dict | list | None",
    id_to_label: dict[str, str],
    answer_id_to_code: dict[int, tuple[str, int]],
    row_answer_id_to_info: dict[int, tuple[str, str]],
    q_ids_with_answer_ids: "set[str] | None" = None,
) -> "dict | list | None":
    """Recursively translate a show_condition / contradiction rule node.

    Supported id formats
    --------------------
    ``"799409"``
        Plain question_id — SA / MA / ranking / NUM.
        If the question has answer_ids in its choices: translate value → codes.
        If not (e.g. area/special types): treat value as direct choice code.

    ``"799010-5582321"``
        Compound id ``"{question_id}-{row_answer_id}"`` — matrix sub-question.

    Raises
    ------
    ValueError
        When a question_id or row_answer_id cannot be resolved,
        or when answer_id lookup fails for a question that has answer_ids.
    """
    if node is None:
        return None

    if isinstance(node, list):
        return [
            _translate_condition_node(
                r, id_to_label, answer_id_to_code, row_answer_id_to_info, q_ids_with_answer_ids
            )
            for r in node
        ]

    if "condition" in node:
        out: dict = {
            "condition": node["condition"],
            "rules": [
                _translate_condition_node(
                    r, id_to_label, answer_id_to_code, row_answer_id_to_info, q_ids_with_answer_ids
                )
                for r in node.get("rules", [])
            ],
        }
        if "ignore_no_data" in node:
            out["ignore_no_data"] = node["ignore_no_data"]
        return out

    # ── Leaf rule ──────────────────────────────────────────────────────────
    q_id_str   = str(node.get("id", ""))
    operator   = node.get("operator", "")
    input_type = node.get("input", "select")
    raw_value  = node.get("value")

    # ── Matrix compound id: "{question_id}-{row_answer_id}" ───────────────
    if "-" in q_id_str:
        q_id_part, row_aid_str = q_id_str.split("-", 1)
        q_label = id_to_label.get(q_id_part)
        if q_label is None:
            raise ValueError(
                f"show_condition/contradiction rule: question_id '{q_id_part}' "
                "(from compound id '{q_id_str}') not found in survey definition"
            )
        row_info = row_answer_id_to_info.get(int(row_aid_str))
        if row_info is None:
            # Fallback: some QMe surveys encode MA sub-choice conditions as
            # "{q_id}-{answer_id}" where the second part is a flat choice answer_id
            # (not a matrix row). Resolve via answer_id_to_code.
            flat_info = answer_id_to_code.get(int(row_aid_str))
            if flat_info is not None:
                _, choice_code = flat_info
                return {
                    "question":       q_label,
                    "operator":       _normalize_operator(operator),
                    "ignore_no_data": node.get("ignore_no_data", 0),
                    "codes":          [choice_code],
                }
            # Cannot resolve — log and return None so the rule is skipped
            # (eval_condition_vec treats None as True = no constraint)
            logger.warning(
                "show_condition: compound id '%s' — row_answer_id %s not found "
                "in matrix rows or flat choices; rule will be skipped",
                q_id_str, row_aid_str,
            )
            return None
        _, row_code = row_info
        sub_q_label = f"{q_label}_r{row_code}"

        translated: dict = {
            "question":        sub_q_label,
            "operator":        _normalize_operator(operator),
            "ignore_no_data":  node.get("ignore_no_data", 0),
        }
        if input_type == "select" and raw_value is not None and not _is_numeric_op(operator) and operator not in _NULL_OPS:
            # value is a column answer_id (may arrive as string or int or list)
            if isinstance(raw_value, list):
                codes = []
                for aid in raw_value:
                    info = answer_id_to_code.get(int(aid))
                    if info is None:
                        raise ValueError(
                            f"column answer_id {aid} (matrix '{sub_q_label}') "
                            "not found in any choice"
                        )
                    codes.append(info[1])
                translated["codes"] = codes
            else:
                info = answer_id_to_code.get(int(raw_value))
                if info is None:
                    raise ValueError(
                        f"column answer_id {raw_value} (matrix '{sub_q_label}') "
                        "not found in any choice"
                    )
                translated["codes"] = [info[1]]
        else:
            translated["value"] = raw_value
        return translated

    # ── Plain question_id ─────────────────────────────────────────────────
    q_label = id_to_label.get(q_id_str)
    if q_label is None:
        raise ValueError(
            f"show_condition/contradiction rule: question_id '{q_id_str}' "
            "not found in survey definition"
        )

    translated = {
        "question":       q_label,
        "operator":       _normalize_operator(operator),
        "ignore_no_data": node.get("ignore_no_data", 0),
    }

    # Questions without answer_ids in their choices (e.g. area/special types)
    # use direct choice codes in condition rules — no translation needed.
    _needs_translation = (q_ids_with_answer_ids is None) or (q_id_str in q_ids_with_answer_ids)

    if input_type == "select" and not _is_numeric_op(operator) and operator not in _NULL_OPS:
        if isinstance(raw_value, list):
            codes = []
            for aid in raw_value:
                vi = int(aid)
                if _needs_translation:
                    info = answer_id_to_code.get(vi)
                    if info is None:
                        raise ValueError(
                            f"answer_id {aid} (question '{q_label}') not found in any choice"
                        )
                    codes.append(info[1])
                else:
                    codes.append(vi)   # direct code
            translated["codes"] = codes
        elif raw_value is not None:
            vi = int(raw_value)
            if _needs_translation:
                info = answer_id_to_code.get(vi)
                if info is None:
                    raise ValueError(
                        f"answer_id {raw_value} (question '{q_label}') not found in any choice"
                    )
                translated["codes"] = [info[1]]
            else:
                translated["codes"] = [vi]   # direct code
    else:
        if raw_value is not None:
            translated["value"] = raw_value

    return translated


# ── Operator helpers (used by _translate_condition_node) ──────────────────────

# Operators that indicate "count of answers" comparison — value is a number, not answer_id.
# QMe uses these even when input="select".
_NUMBER_ANSWER_PREFIX = "number_answer_"

_NUMBER_ANSWER_MAP: dict[str, str] = {
    "number_answer_equal":              "count_equal",
    "number_answer_not_equal":          "count_not_equal",
    "number_answer_less":               "count_less",
    "number_answer_less_or_equal":      "count_less_or_equal",
    "number_answer_greater":            "count_greater",
    "number_answer_greater_or_equal":   "count_greater_or_equal",
}

# Null-check operators — value is always null, no code translation needed.
_NULL_OPS = {"is_null", "is_not_null"}


def _is_numeric_op(op: str) -> bool:
    return op.startswith(_NUMBER_ANSWER_PREFIX)


def _normalize_operator(op: str) -> str:
    """Normalize QMe-specific operators to a standard internal name."""
    return _NUMBER_ANSWER_MAP.get(op, op)


def _detect_other_codes(choices_i18n: dict) -> list[str]:
    """Return choice codes that are "other-specify" inputs (is_other: true)."""
    return [
        str(code)
        for code, i18n in choices_i18n.items()
        if isinstance(i18n, dict) and i18n.get("is_other")
    ]


_TYPE_NAME_LABEL: dict[str, str] = {
    "gender":          "SA",
    "married-status":  "SA",
    "area":            "SA",
    "personal-income": "SA",
    "photo":           "photo",
}

# ── Synthetic choices for special question types ──────────────────────────────

GENDER_CHOICES: dict[str, dict] = {
    "1": {"vi": "Nam",  "en": "Male"},
    "2": {"vi": "Nữ",   "en": "Female"},
}

# Fixed city/province codes for area questions (QMe code list — 63 provinces).
# Unknown values are auto-assigned 9000, 9001, 9002, … to avoid conflicts.
AREA_BASE_CHOICES: dict[str, dict] = {
    "10":   {"vi": "Cần Thơ",            "en": "Can Tho"},
    "11":   {"vi": "Hồ Chí Minh",        "en": "Ho Chi Minh City"},
    "12":   {"vi": "Hà Nội",             "en": "Hanoi"},
    "13":   {"vi": "Hải Phòng",          "en": "Hai Phong"},
    "14":   {"vi": "Đà Nẵng",            "en": "Da Nang"},
    "15":   {"vi": "Đồng Nai",           "en": "Dong Nai"},
    "16":   {"vi": "Bình Dương",         "en": "Binh Duong"},
    "17":   {"vi": "Long An",            "en": "Long An"},
    "18":   {"vi": "An Giang",           "en": "An Giang"},
    "19":   {"vi": "Bà Rịa - Vũng Tàu", "en": "Ba Ria - Vung Tau"},
    "20":   {"vi": "Bắc Giang",          "en": "Bac Giang"},
    "21":   {"vi": "Bắc Kạn",            "en": "Bac Kan"},
    "22":   {"vi": "Bạc Liêu",           "en": "Bac Lieu"},
    "23":   {"vi": "Bắc Ninh",           "en": "Bac Ninh"},
    "24":   {"vi": "Bến Tre",            "en": "Ben Tre"},
    "25":   {"vi": "Bình Định",          "en": "Binh Dinh"},
    "26":   {"vi": "Bình Phước",         "en": "Binh Phuoc"},
    "27":   {"vi": "Bình Thuận",         "en": "Binh Thuan"},
    "28":   {"vi": "Cà Mau",             "en": "Ca Mau"},
    "29":   {"vi": "Cao Bằng",           "en": "Cao Bang"},
    "30":   {"vi": "Đắk Lắk",           "en": "Dak Lak"},
    "31":   {"vi": "Đắk Nông",           "en": "Dak Nong"},
    "32":   {"vi": "Điện Biên",          "en": "Dien Bien"},
    "33":   {"vi": "Đồng Tháp",          "en": "Dong Thap"},
    "34":   {"vi": "Gia Lai",            "en": "Gia Lai"},
    "35":   {"vi": "Hà Nam",             "en": "Ha Nam"},
    "36":   {"vi": "Hà Giang",           "en": "Ha Giang"},
    "37":   {"vi": "Hà Tĩnh",            "en": "Ha Tinh"},
    "38":   {"vi": "Hải Dương",          "en": "Hai Duong"},
    "39":   {"vi": "Hậu Giang",          "en": "Hau Giang"},
    "40":   {"vi": "Hòa Bình",           "en": "Hoa Binh"},
    "41":   {"vi": "Hưng Yên",           "en": "Hung Yen"},
    "42":   {"vi": "Khánh Hòa",          "en": "Khanh Hoa"},
    "43":   {"vi": "Kiên Giang",         "en": "Kien Giang"},
    "44":   {"vi": "Kon Tum",            "en": "Kon Tum"},
    "45":   {"vi": "Lai Châu",           "en": "Lai Chau"},
    "46":   {"vi": "Lâm Đồng",           "en": "Lam Dong"},
    "47":   {"vi": "Lạng Sơn",           "en": "Lang Son"},
    "48":   {"vi": "Lào Cai",            "en": "Lao Cai"},
    "49":   {"vi": "Nam Định",           "en": "Nam Dinh"},
    "50":   {"vi": "Nghệ An",            "en": "Nghe An"},
    "51":   {"vi": "Ninh Bình",          "en": "Ninh Binh"},
    "52":   {"vi": "Ninh Thuận",         "en": "Ninh Thuan"},
    "53":   {"vi": "Phú Thọ",            "en": "Phu Tho"},
    "54":   {"vi": "Phú Yên",            "en": "Phu Yen"},
    "55":   {"vi": "Quảng Bình",         "en": "Quang Binh"},
    "56":   {"vi": "Quảng Nam",          "en": "Quang Nam"},
    "57":   {"vi": "Quảng Ninh",         "en": "Quang Ninh"},
    "58":   {"vi": "Quảng Trị",          "en": "Quang Tri"},
    "59":   {"vi": "Sóc Trăng",          "en": "Soc Trang"},
    "60":   {"vi": "Sơn La",             "en": "Son La"},
    "61":   {"vi": "Tây Ninh",           "en": "Tay Ninh"},
    "62":   {"vi": "Thái Bình",          "en": "Thai Binh"},
    "63":   {"vi": "Thái Nguyên",        "en": "Thai Nguyen"},
    "64":   {"vi": "Thanh Hóa",          "en": "Thanh Hoa"},
    "65":   {"vi": "Tiền Giang",         "en": "Tien Giang"},
    "66":   {"vi": "Trà Vinh",           "en": "Tra Vinh"},
    "67":   {"vi": "Tuyên Quang",        "en": "Tuyen Quang"},
    "68":   {"vi": "Vĩnh Long",          "en": "Vinh Long"},
    "69":   {"vi": "Vĩnh Phúc",          "en": "Vinh Phuc"},
    "70":   {"vi": "Yên Bái",            "en": "Yen Bai"},
    "71":   {"vi": "Quảng Ngãi",         "en": "Quang Ngai"},
    "2109": {"vi": "Thừa Thiên - Huế",  "en": "Thua Thien - Hue"},
}

# personal-income choices are auto-discovered at encode time (actual answer text
# from QMe uses VND number ranges like "10,000,001 - 15,000,000 VND").
# Codes are assigned by sorting ranges by their lower bound (lowest = code 1).
# "Don't know" variants are always assigned code 99.
# The filled-in choices_i18n is persisted back into metadata.json after encode.
PERSONAL_INCOME_CHOICES: dict[str, dict] = {}   # placeholder — populated at runtime

def _resolve_answer_type(q_type: int, input_type: int, type_name: str = "") -> str:
    if type_name in _TYPE_NAME_LABEL:
        return _TYPE_NAME_LABEL[type_name]
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
          "title_i18n":   null,              ← AI-summarized title; null until Claude fills it
                                                in post-ingestion (see CLAUDE.md Step 3a) — this
                                                function never generates real text for it, only
                                                the null placeholder (deterministic code has no
                                                summarization capability of its own).
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

    # ── Build lookup maps for show_condition / contradiction translation ───
    _all_qs = definition.get("questions", [])
    _id_to_label, _answer_id_to_code, _row_aid_to_info, _q_ids_with_aids = _build_lookup_maps(_all_qs)

    questions: dict[str, dict] = {}
    for q in _all_qs:
        q_type   = q["type"]
        inp_type = q.get("input_type", 0)
        atype    = _resolve_answer_type(q_type, inp_type, q.get("type_name", ""))

        if atype in _EX:
            continue

        raw_choices  = q.get("choices", {})
        choices_i18n = q.get("choices_i18n", {})
        pos          = q["position"]
        qid          = q["question_id"]
        col          = f"q{pos}"
        key          = str(qid)

        # ── matrix: build sub_questions (one per row) ────────────────────────
        # rows/columns can be in `choices` (old format) or `choices_i18n` (new format)
        _ci_rows = choices_i18n.get("rows", {}) if isinstance(choices_i18n, dict) else {}
        _raw_rows = (raw_choices.get("rows") if isinstance(raw_choices, dict) else None) or _ci_rows

        if atype in _MATRIX_SUB_TYPE and _raw_rows:
            def _row_label(v: object) -> str:
                if isinstance(v, dict):
                    return v.get("vi") or v.get("en") or str(v)
                return str(v)

            rows_map    = {str(k): _row_label(v) for k, v in _raw_rows.items()}
            cols_i18n   = {str(k): v for k, v in choices_i18n.get("columns", {}).items()} if isinstance(choices_i18n, dict) else {}
            sub_atype   = _MATRIX_SUB_TYPE[atype]

            # detect "other" column codes once (shared across all rows)
            matrix_other_codes = _detect_other_codes(cols_i18n) if cols_i18n else []

            parent_label = q.get("label") or col
            sub_questions: dict[str, dict] = {}
            for row_code, row_label in rows_map.items():
                sub_key      = f"{col}_r{row_code}"
                sub_col_name = f"{parent_label}_r{row_code}"   # matches rawdata column name
                sub_entry: dict = {
                    "parent":       key,
                    "row_index":    row_code,
                    "answer_type":  sub_atype,
                    "label":        sub_col_name,   # = rawdata column name (e.g. Q9_1_r1)
                    "row_label":    row_label,       # choice text (e.g. "Wakodo")
                    "title_i18n":   None,            # AI-summarized title (vi/en); see CLAUDE.md
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

        # ── synthetic choices for special types ─────────────────────────────
        import copy as _copy
        type_name_str = q.get("type_name", "")
        synthetic_type = ""
        if type_name_str == "gender":
            choices_i18n  = _copy.deepcopy(GENDER_CHOICES)
            synthetic_type = "gender"
        elif type_name_str == "area" or q_type == 1109:
            choices_i18n  = _copy.deepcopy(AREA_BASE_CHOICES)
            synthetic_type = "area"
        elif type_name_str == "personal-income":
            choices_i18n  = _copy.deepcopy(PERSONAL_INCOME_CHOICES)
            synthetic_type = "personal-income"

        q_entry: dict = {
            "position":     pos,
            "question_id":  qid,
            "label":        q.get("label", ""),
            "question_i18n": q.get("question_i18n", {}),
            "title_i18n":   None,   # AI-summarized title (vi/en), filled by Claude post-ingestion; see CLAUDE.md
            "answer_type":  atype,
            "mandatory":    q.get("mandatory", False),
            "status":       q.get("status", 1),
            "choices_i18n": choices_i18n,
        }
        if synthetic_type:
            q_entry["synthetic_choices"] = True
            q_entry["synthetic_type"]    = synthetic_type
        if other_codes:
            q_entry["other_choice_codes"] = other_codes
        if sub_questions:
            q_entry["children"]      = list(sub_questions.keys())
            q_entry["sub_questions"] = sub_questions

        # ── show_condition + contradiction_settings (translated) ───────────
        raw_sc = q.get("show_condition")
        raw_cs = q.get("contradiction_settings")
        if raw_sc:
            q_entry["show_condition"] = _translate_condition_node(
                raw_sc, _id_to_label, _answer_id_to_code, _row_aid_to_info, _q_ids_with_aids
            )
        if raw_cs:
            raw_rules = raw_cs.get("rules")
            q_entry["contradiction_settings"] = {
                "action": raw_cs.get("action"),
                "rules":  _translate_condition_node(
                    raw_rules, _id_to_label, _answer_id_to_code, _row_aid_to_info, _q_ids_with_aids
                ) if raw_rules else None,
            }

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
