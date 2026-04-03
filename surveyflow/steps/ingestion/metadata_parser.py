"""Parse Qme survey definition into metadata.json structure."""

from __future__ import annotations

# Question types that carry no respondent answer (instructions / section headers)
_INSTRUCTION_TYPES = {31}

# Map definition question type → human-readable answer_type string
_TYPE_LABEL: dict[int, str] = {
    1:    "freetext",
    2:    "singlechoice",
    3:    "multiplechoice",
    4:    "matrix",
    6:    "ranking",
    8:    "singlechoice",      # piped single choice
    9:    "multiplechoice",    # unaided / piped MA
    31:   "instruction",
    40:   "audio",
    1106: "user-name",
    1107: "user-phone",
    1109: "area",
}

# input_type overrides for type=1 (freetext)
_INPUT_TYPE_LABEL: dict[int, str] = {
    3:   "singlenumber",
    68:  "reward",
    100: "multiplenumber",
}


def _resolve_answer_type(q_type: int, input_type: int) -> str:
    if q_type == 1 and input_type in _INPUT_TYPE_LABEL:
        return _INPUT_TYPE_LABEL[input_type]
    return _TYPE_LABEL.get(q_type, f"unknown_{q_type}")


def parse_metadata(definition: dict) -> dict:
    """Convert a raw ``get_survey_definition`` response to metadata dict.

    Parameters
    ----------
    definition : dict
        Raw JSON response from the Qme ``get_survey_definition`` tool.

    Returns
    -------
    dict
        Metadata dict.  Saved to ``metadata.json`` by the IO layer.
        ``questions[*].values`` starts empty and is populated later by
        :func:`enrich_metadata_values` once all rows have been parsed.
    """
    survey = definition["survey"]

    questions = []
    for q in definition.get("questions", []):
        q_type    = q["type"]
        inp_type  = q.get("input_type", 0)
        is_instr  = q_type in _INSTRUCTION_TYPES

        questions.append({
            "position":         q["position"],
            "label":            f"q{q['position']}",
            "question_id":      q["question_id"],
            "question":         q["question"],
            "english_question": q["english_question"],
            "type":             q_type,
            "input_type":       inp_type,
            "answer_type":      _resolve_answer_type(q_type, inp_type),
            "is_instruction":   is_instr,
            "mandatory":        q.get("mandatory", False),
            "status":           q.get("status", 1),
            # populated by enrich_metadata_values()
            "values":           {},
        })

    return {
        "survey_id":      survey["survey_id"],
        "title":          survey.get("title", ""),
        "english_title":  survey.get("english_title", ""),
        "status":         survey.get("status", ""),
        "start_date":     survey.get("start_date", ""),
        "end_date":       survey.get("end_date", ""),
        "question_count": len(questions),
        "questions":      questions,
    }


def enrich_metadata_values(metadata: dict, rawdata_records: list[dict]) -> dict:
    """Populate ``values`` for each question from observed row answers.

    For singlechoice / multiplechoice / ranking questions the ``values``
    dict maps each answer label to an auto-assigned integer code
    (1-based, in order of first appearance).

    Parameters
    ----------
    metadata : dict
        Metadata dict produced by :func:`parse_metadata`.
    rawdata_records : list[dict]
        List of flat dicts produced by ``data_parser.parse_rows``,
        one per respondent row (before DataFrame conversion).

    Returns
    -------
    dict
        The same ``metadata`` dict, mutated in place, then returned.
    """
    # Build label → question meta quick-lookup
    label_to_meta = {q["label"]: q for q in metadata["questions"]}

    # Collect ordered unique values per label
    seen: dict[str, list[str]] = {}

    for record in rawdata_records:
        for label, raw_value in record.items():
            if label not in label_to_meta or not raw_value:
                continue
            q = label_to_meta[label]
            atype = q["answer_type"]

            if atype in ("singlechoice",):
                vals = [str(raw_value).strip()]
            elif atype in ("multiplechoice", "ranking"):
                vals = [v.strip() for v in str(raw_value).split(";") if v.strip()]
            else:
                continue  # freetext / number / photo / etc. — no code mapping

            bucket = seen.setdefault(label, [])
            for v in vals:
                if v not in bucket:
                    bucket.append(v)

    # Write back as {answer_label: code}
    for label, values in seen.items():
        label_to_meta[label]["values"] = {v: i + 1 for i, v in enumerate(values)}

    return metadata
