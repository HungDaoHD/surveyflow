# SurveyFlow — Claude Workflow Guide

This project uses the **surveyflow** Python package to process survey data from QMe into:
`rawdata.csv` + `metadata.json` + `datatable.xlsx`

---

## Core concept

`datatable.json` is the **single source of truth** for what the user wants in the output table.
- Claude modifies `datatable.json` based on user requests (natural language)
- Then runs the pipeline to produce updated `datatable.xlsx`
- Users never edit `datatable.json` directly — they tell Claude what they want

---

## Environment setup (check once at start of every session)

Before doing anything, verify the environment is ready:

```bash
# 1. Check surveyflow installed
python -c "import surveyflow; print(surveyflow.__version__)"
```

If the command fails → install it:
```bash
pip install surveyflow
```

If running from a local dev folder (editable install):
```bash
pip install -e .
```

---

## Workflow A — First time running a survey

### Step 1 — Find survey
```
search_surveys(query="SURVEY_NAME")
```
Note the `survey_id`.

### Step 2 — Fetch and save MCP data
**First, check if `output/SURVEY_NAME/input/` already exists and contains `definition.json` + `rows_page_*.json`.**

If files exist → ask user:
> "Input data đã có sẵn (`definition.json`, `rows_page_1.json`, …). Bạn muốn dùng data cũ hay fetch lại data mới từ QMe?"

- User says **dùng data cũ / no** → skip to Step 3
- User says **fetch lại / yes** → proceed to fetch below

If files do not exist → fetch immediately (no need to ask):
```
get_survey_definition(survey_id)        →  save to output/SURVEY_NAME/input/definition.json
get_survey_rows(survey_id, offset=0)    →  save to output/SURVEY_NAME/input/rows_page_1.json
get_survey_rows(survey_id, offset=200)  →  save to output/SURVEY_NAME/input/rows_page_2.json
... keep fetching until has_more = false
```

> ⚠️ **MCP trả về structuredContent (dict), không phải plain text.**
> Dùng Write tool để ghi file — nội dung là `json.dumps(result, ensure_ascii=False, indent=2)`.
> Không dùng bash pipe cho MCP output.

### Step 3 — Create datatable.json
Create `output/SURVEY_NAME/datatable.json` based on:
- The survey definition (question positions, types, labels)
- User's request (if specified)
- If user does not specify → auto-generate (see Auto-generate rules below)

### Step 4 — Run pipeline

**Option A** — nếu `run_pipeline.py` có sẵn trong project root:
```bash
python run_pipeline.py \
  --input-dir        output/SURVEY_NAME/input \
  --output-dir       output/SURVEY_NAME \
  --version          v1 \
  --datatable-config output/SURVEY_NAME/datatable.json
```

**Option B** — nếu không có `run_pipeline.py` (fresh environment), chạy inline:
```python
import json
from pathlib import Path
from surveyflow import Pipeline, PipelineConfig

input_dir  = Path("output/SURVEY_NAME/input")
output_dir = Path("output/SURVEY_NAME")

with open(input_dir / "definition.json", encoding="utf-8") as f:
    definition = json.load(f)

rows_pages = [
    json.load(open(p, encoding="utf-8"))
    for p in sorted(input_dir.glob("rows_page_*.json"))
]

result = Pipeline(PipelineConfig(
    definition       = definition,
    rows_pages       = rows_pages,
    output_dir       = str(output_dir),
    version          = "v1",
    datatable_config = "output/SURVEY_NAME/datatable.json",
)).run()

print(result["datatable_path"])
```

---

## Workflow B — User requests changes to the table

When user says things like:
- *"thêm income vào banner"*
- *"bỏ Q15 ra khỏi stub"*
- *"thêm mean và std cho Q36"*
- *"đổi banner gender thành Male/Female/Other"*
- *"thêm T2B cho tất cả câu singlechoice"*

**Always:**
1. Read current `output/SURVEY_NAME/datatable.json`
2. Modify it according to user request
3. Save `datatable.json`
4. Run pipeline with new version (Option A or B above, increment version v1→v2→v3…)

---

## datatable.json structure

```json
{
  "title": "SURVEY_NAME - Data Table",
  "significance_test": {
    "enabled": true,
    "levels": [90, 95],
    "method": "independent"
  },
  "banner": [
    { "label": "Total", "filter": null },
    {
      "label": "Gender",
      "question": "Q10",
      "groups": [
        { "label": "Male",   "value": 2 },
        { "label": "Female", "value": 1 }
      ]
    },
    {
      "label": "Age Group",
      "question": "Q12",
      "groups": [
        { "label": "Under 30", "values": [3, 4, 5] },
        { "label": "30 - 39",  "values": [1, 2]    },
        { "label": "40+",      "values": [6, 7, 8]  }
      ]
    }
  ],
  "stub": [
    { "question": "Q10", "label": "Gender",     "stats": ["base", "percent"] },
    { "question": "Q36", "label": "Food Freq",  "stats": ["base", "percent", "t2b", "b2b", "mean", "std", "se"] }
  ],
  "tables": [
    { "sheet": "Count",            "cell_content": "count",      "show_sig": false },
    { "sheet": "Percentage",       "cell_content": "percentage", "show_sig": false },
    { "sheet": "Percentage & Sig", "cell_content": "percentage", "show_sig": true  }
  ]
}
```

### Banner rules
- Always include `{ "label": "Total", "filter": null }` as first entry
- `value` = single integer code from rawdata.csv
- `values` (plural) = list of codes to group together (e.g. age ranges)
- **Question reference: use the question label** (e.g. `"Q10"`) — this is the `label` field in metadata.json
  - `q{position}` format (e.g. `"q10"`) is also accepted for backward compatibility
- To find choice codes → read `output/SURVEY_NAME/vX/metadata.json` → find the question by label → inspect `choices_i18n`
- **MA questions are supported as banner** — each code becomes a column; respondents can appear in multiple columns

### Stub rules
- One entry per question
- `stats` options: `"base"`, `"percent"`, `"t2b"`, `"b2b"`, `"mean"`, `"std"`, `"se"`
- Supported answer types: `SA` (singlechoice), `MA` (multiplechoice), `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`
  - Matrix questions automatically expand into one block per row (sub-question)
- Excluded automatically: `FT`, `NUM`, `instruction`, `user-name`, `user-phone`, `reward`, `record`
- **Stub order follows datatable.json** — the pipeline outputs questions in the order they appear in `stub`
  - Auto-generate sorts by position; after that, Claude must preserve the user's order

### Tables rules
- Always keep all 3 sheets: Count, Percentage, Percentage & Sig
- Do not modify `tables` unless user explicitly requests it

---

## Auto-generate datatable.json rules
When user does not specify — use this logic:

**Banner:** Pick singlechoice (SA) questions that look demographic:
- Gender, Age, Location, Income, Marital status, Occupation
- Max 4-5 banner groups + Total

**Stub:** Include all codeable questions sorted by position:
- `SA` (singlechoice), `MA` (multiplechoice), `Matrix_SA`, `Matrix_MA`
- Default stats: `["base", "percent"]`
- Skip: `FT`, `NUM`, `instruction`, `user-name`, `user-phone`, `reward`, `record`

---

## Output structure
```
output/SURVEY_NAME/
├── input/
│   ├── definition.json       ← from get_survey_definition (fetch once)
│   ├── rows_page_1.json      ← from get_survey_rows (fetch once)
│   └── rows_page_2.json
├── datatable.json            ← Claude manages this file
├── v1/                       ← first run
│   ├── rawdata.csv
│   ├── metadata.json
│   └── datatable.xlsx
├── v2/                       ← after user requests changes
│   └── datatable.xlsx
└── v3/
    └── datatable.xlsx
```

> **Note:** `input/` data is fetched once and reused across all versions.
> **Always ask before re-fetching:** "Input data đã có sẵn (`definition.json`, `rows_page_*.json`). Bạn muốn dùng data cũ hay fetch lại data mới từ QMe?"
> Only re-fetch if user confirms yes.

---

## Common user requests → datatable.json changes

| User says | Claude does |
|---|---|
| "thêm income vào banner" | Add income question to `banner` array |
| "bỏ Q15 khỏi stub" | Remove Q15 entry from `stub` array |
| "thêm mean std cho Q36" | Add `"mean"`, `"std"` to Q36's `stats` |
| "thêm tất cả câu vào stub" | Add all codeable questions to `stub` |
| "tắt sig test" | Set `significance_test.enabled = false` |
| "chỉ chạy 1 sheet percentage" | Modify `tables` to keep only Percentage sheet |
| "refresh data / lấy data mới" | Re-fetch MCP rows → overwrite `input/rows_page_*.json` → re-run |

---

## Profile status
Default: `approved` only.
Override: `--profile-status approved,pending`
