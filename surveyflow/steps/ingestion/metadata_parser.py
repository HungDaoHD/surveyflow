"""Parse Qme survey definition into metadata.json structure."""

from __future__ import annotations

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

    The definition now includes a ``choices`` field per codeable question
    (singlechoice / multiplechoice / ranking) containing a canonical
    ``{code_str: label}`` mapping supplied by the server.  This replaces
    the old first-appearance heuristic and makes metadata complete even
    for answer options that were never selected.

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
          "values": {"1": "Male", "2": "Female"}   ← from choices
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

        # choices = {code_str: label} supplied by server for codeable questions
        raw_choices = q.get("choices", {})
        values = {str(k): str(v) for k, v in raw_choices.items()}

        label = f"q{q['position']}"
        questions[label] = {
            "position":         q["position"],
            "question_id":      q["question_id"],
            "question":         q["question"],
            "english_question": q["english_question"],
            "answer_type":      atype,
            "mandatory":        q.get("mandatory", False),
            "status":           q.get("status", 1),
            "values":           values,
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
