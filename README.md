# surveyflow

A Python library for processing survey data — parse survey definitions and responses into structured outputs ready for analysis.

## Features

- Parse survey **definition** (question structure, types, positions) into `metadata.json`
- Parse survey **response rows** into `rawdata.csv` with numeric codes
  - Single-choice → integer code (e.g. `1`)
  - Multi-choice / ranking → semicolon-separated codes (e.g. `"1;3;5"`)
  - Open-ended / matrix / number → raw text
- Filter responses by status (default: `approved` only)
- Consistent columns between `rawdata.csv` and `metadata.json`

## Installation

```bash
pip install surveyflow
```

## Quick Start

```python
from surveyflow.steps.ingestion import IngestionStep

# definition: dict from your survey platform's definition API
# rows_pages: list of paginated response pages from your survey platform

step = IngestionStep()
context = step.run({
    "definition":  definition,
    "rows_pages":  rows_pages,
    "output_dir":  "./output",
})

df       = context["rawdata"]      # pandas DataFrame
metadata = context["metadata"]     # dict with question info + value labels
```

## Output

### `rawdata.csv`

| task_id | date_time | q6 | q7 | q10 | q18 |
|---|---|---|---|---|---|
| task_001 | 2026-03-01 | 1 | 2 | 1 | 1;3 |
| task_002 | 2026-03-01 | 2 | 1 | 2 | 2 |

### `metadata.json`

```json
{
  "survey_id": 12345,
  "questions": {
    "q6": {
      "position": 6,
      "english_question": "Please provide your current address",
      "answer_type": "singlechoice",
      "values": { "1": "Ward 1", "2": "Ward 2", "3": "Ward 3" }
    },
    "q18": {
      "position": 18,
      "english_question": "Who do you live with",
      "answer_type": "multiplechoice",
      "values": { "1": "Spouse", "2": "Parents", "3": "Children" }
    }
  }
}
```

## Input Format

### `definition`

```python
{
    "survey": { "survey_id": 12345, "title": "...", ... },
    "questions": [
        {
            "question_id": 1001,
            "position": 6,
            "question": "...",
            "english_question": "Please provide your current address",
            "type": 2,        # 2=singlechoice, 3=multiplechoice, 6=ranking, 4=matrix, ...
            "input_type": 0,
            "mandatory": True,
            "status": 1
        },
        ...
    ]
}
```

### `rows_pages`

```python
[
    {   # page 1
        "rows": [
            {
                "task_id": "task_001",
                "date_time": "2026-03-01 09:00:00",
                "profile_status": "approved",
                "questions": [
                    { "type": "singlechoice", "question": "Please provide your current address", "answer": "Ward 1" },
                    { "type": "multiplechoice", "question": "Who do you live with",
                      "answer": [{"answer_name": "Spouse"}, {"answer_name": "Children"}] },
                    ...
                ]
            },
            ...
        ]
    },
    # page 2, page 3, ...
]
```

## Answer Types

| `type` value | `answer_type` | Encoded in rawdata? |
|---|---|---|
| 2 | `singlechoice` | Yes → int |
| 3 | `multiplechoice` | Yes → `"1;3;5"` |
| 6 | `ranking` | Yes → `"2;1;3"` |
| 4 | `matrix` | No → `"row:col\|row:col"` |
| 1 + input_type=100 | `multiplenumber` | No → `"label:num\|label:num"` |
| 1 | `freetext` | No → raw text |
| 1 + input_type=3 | `singlenumber` | No → raw number |
| 1109 | `area` | No → raw text |

Excluded from output: `audio`, `user-name`, `user-phone`, `instruction`, `reward`.

## Profile Status Filter

```python
# Default: approved only
step.run({ ..., "profile_status": ["approved"] })

# Include all statuses
step.run({ ..., "profile_status": [] })

# Custom filter
step.run({ ..., "profile_status": ["approved", "pending"] })
```

## Requirements

- Python >= 3.10
- pandas >= 2.0
- openpyxl >= 3.1
