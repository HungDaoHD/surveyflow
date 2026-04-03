"""Parse Qme survey definition into metadata.json structure."""

from __future__ import annotations

_INSTRUCTION_TYPES = {31}

_TYPE_LABEL: dict[int, str] = {
    1:    "freetext",
    2:    "singlechoice",
    3:    "multiplechoice",
    4:    "matrix",
    6:    "ranking",
    8:    "singlechoice",
    9:    "multiplechoice",
    31:   "instruction",
    40:   "audio",
    1106: "user-name",
    1107: "user-phone",
    1109: "area",
}

_TYPE_LABEL: dict[int, str] = {
    1:    "freetext",
    2:    "singlechoice",
    3:    "multiplechoice",
    4:    "matrix",
    6:    "ranking",
    1109: "area",
}


_INPUT_TYPE_LABEL: dict[int, str] = {
    3:   "singlenumber",
    68:  "reward",
    100: "multiplenumber",
}

# answer_types that support numeric coding
CODEABLE_TYPES = {"singlechoice", "multiplechoice", "ranking"}


def _resolve_answer_type(q_type: int, input_type: int) -> str:
    if q_type == 1 and input_type in _INPUT_TYPE_LABEL:
        return _INPUT_TYPE_LABEL[input_type]
    return _TYPE_LABEL.get(q_type, f"unknown_{q_type}")


def parse_metadata(definition: dict) -> dict:
    """Convert get_survey_definition response → metadata dict.

    Excluded answer_types (audio, user-name, user-phone, instruction, reward)
    are omitted — kept in sync with rawdata.csv.

    Structure:
    {
      "survey_id": ...,
      "questions": {
        "q7": {
          "position": 7,
          "question_id": ...,
          "question": "...(VN)...",
          "english_question": "...",
          "answer_type": "singlechoice",
          "mandatory": true,
          "status": 1,
          "values": {}     ← populated by enrich_metadata_values()
        },
        ...
      }
    }
    """
    from surveyflow.steps.ingestion.data_parser import EXCLUDED_ANSWER_TYPES

    survey = definition["survey"]

    questions: dict[str, dict] = {}
    for q in definition.get("questions", []):
        q_type   = q["type"]
        inp_type = q.get("input_type", 0)
        atype    = _resolve_answer_type(q_type, inp_type)

        if atype in EXCLUDED_ANSWER_TYPES:
            continue

        label = f"q{q['position']}"
        questions[label] = {
            "position":         q["position"],
            "question_id":      q["question_id"],
            "question":         q["question"],
            "english_question": q["english_question"],
            "answer_type":      atype,
            "mandatory":        q.get("mandatory", False),
            "status":           q.get("status", 1),
            "values":           {},   # {code_str: label_text}
        }

    return {
        "survey_id":     survey["survey_id"],
        "title":         survey.get("title", ""),
        "english_title": survey.get("english_title", ""),
        "status":        survey.get("status", ""),
        "start_date":    survey.get("start_date", ""),
        "end_date":      survey.get("end_date", ""),
        "questions":     questions,
    }


def build_encoding_map(
    raw_text_records: list[dict],
    metadata: dict,
) -> dict[str, dict[str, int]]:
    """Build {label: {answer_text: code}} mapping from observed row data.

    Codes are assigned 1-based in order of first appearance.
    Only CODEABLE_TYPES (singlechoice, multiplechoice, ranking) are encoded.
    """
    questions = metadata["questions"]
    seen: dict[str, list[str]] = {}

    for record in raw_text_records:
        for label, raw_value in record.items():
            q = questions.get(label)
            if q is None or not raw_value:
                continue
            if q["answer_type"] not in CODEABLE_TYPES:
                continue

            parts = [v.strip() for v in str(raw_value).split(";") if v.strip()]
            bucket = seen.setdefault(label, [])
            for p in parts:
                if p not in bucket:
                    bucket.append(p)

    return {
        label: {text: (i + 1) for i, text in enumerate(texts)}
        for label, texts in seen.items()
    }


def enrich_metadata_values(
    metadata: dict,
    encoding_map: dict[str, dict[str, int]],
) -> dict:
    """Populate values = {code_str: label_text} for each question."""
    for label, text_to_code in encoding_map.items():
        q = metadata["questions"].get(label)
        if q:
            q["values"] = {str(code): text for text, code in text_to_code.items()}
    return metadata
