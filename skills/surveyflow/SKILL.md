---
name: surveyflow
description: >
  Workflow skill for processing survey data using the surveyflow Python package.
  Use this skill whenever the user mentions running a survey, processing QMe data,
  building a datatable, or working with surveyflow — even if they don't say "surveyflow"
  explicitly. Triggers on phrases like "chạy survey", "tạo datatable", "thêm banner",
  "bỏ câu khỏi stub", "fetch data từ QMe", "chạy pipeline", "chạy bảng", etc.
---

# SurveyFlow Skill

This skill guides Claude through the full surveyflow pipeline:
**Fetch MCP → Ingestion (rawdata + metadata) → Design datatable.json → Table (xlsx)**

## Environment check (once per session)

```bash
python -c "import surveyflow; print(surveyflow.__version__)"
```

If fails → `pip install surveyflow` (or `pip install -e .` for local dev).

---

## Available tools

| File | Purpose |
|---|---|
| `tools/generate_step4_form.py` | Generate Step 4 interactive HTML form from `metadata.json` — show in preview panel so user picks table type / banner / stub in one shot |
| `skills/surveyflow/datatable-editor.html` | Standalone visual editor for `datatable.json` — open in browser, no server needed. Chrome/Edge: save in-place; Firefox: download |
| `skills/surveyflow/scripts/list_questions.py` | List banner-eligible or stub-eligible questions from `metadata.json` |
| `skills/surveyflow/scripts/check_choices.py` | Print choice codes for a question (ready-to-paste banner groups snippet) |

### Preview panel setup (Step 4 form)

`.claude/launch.json` is git-ignored. Create it once per machine if missing:

```json
{
  "version": "0.0.1",
  "configurations": [
    {
      "name": "step4-form",
      "runtimeExecutable": "python",
      "runtimeArgs": ["-m", "http.server", "7891"],
      "port": 7891
    }
  ]
}
```

Then use `preview_start(name="step4-form")` before showing the form.

---

## Output folder structure

```
output/SURVEY_NAME/
├── mcp/                  ← raw MCP files (fetch once)
│   ├── definition.json
│   └── rows_page_*.json
├── data/                 ← rawdata.csv + metadata.json (ingestion output, reused)
├── datatable/
│   └── datatable.json    ← Claude manages this
├── v1/datatable.xlsx
├── v2/datatable.xlsx
└── ...
```

---

## Workflow A — First run

### Step 1 — Find survey
```
search_surveys(query="SURVEY_NAME")
```
Note the `survey_id`.

### Step 2 — Fetch MCP data

Check if `output/SURVEY_NAME/mcp/` already has `definition.json` + `rows_page_*.json`.

- **Files exist** → Ask: "Data đã có sẵn. Dùng data cũ hay fetch lại?"
- **Files missing** → Fetch immediately:

```
get_survey_definition(survey_id)       → save to output/SURVEY_NAME/mcp/definition.json
get_survey_rows(survey_id, offset=0)   → save to output/SURVEY_NAME/mcp/rows_page_1.json
get_survey_rows(survey_id, offset=200) → save to output/SURVEY_NAME/mcp/rows_page_2.json
... keep fetching until has_more = false
```

> MCP returns structuredContent (dict). Use Write tool with `json.dumps(result, ensure_ascii=False, indent=2)`.

### Step 3 — Run ingestion

Check if `output/SURVEY_NAME/data/rawdata.csv` already exists.
- **Exists** → skip this step.
- **Missing** → run ingestion. Since `datatable/datatable.json` doesn't exist yet, the CLI
  automatically skips the table step:

```bash
python run_pipeline.py \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --output-dir output/SURVEY_NAME
```

This generates:
- `data/rawdata.csv`
- `data/metadata.json`  ← use this to look up question labels and choice codes

Use helper scripts to inspect the output:
```bash
# List all banner-eligible questions (SA, MA)
python skills/surveyflow/scripts/list_questions.py output/SURVEY_NAME/data/metadata.json --type banner

# List all stub-eligible questions (SA, MA, Matrix)
python skills/surveyflow/scripts/list_questions.py output/SURVEY_NAME/data/metadata.json --type stub

# Get choice codes for a specific question (ready-to-paste groups snippet)
python skills/surveyflow/scripts/check_choices.py output/SURVEY_NAME/data/metadata.json --question Q10
```

### Step 4 — Ask user and create datatable.json

**If user already specified requirements** → create `datatable/datatable.json` directly.

**If not specified** → use the **preview panel form** (preferred) or ask sequentially in chat.

#### Option A — Preview panel form (recommended)

Generate and show an interactive form in one shot. User selects all 3 options at once:

```bash
python tools/generate_step4_form.py \
  output/SURVEY_NAME/data/metadata.json \
  step4_form.html \
  "SURVEY_NAME"
```

Then start the preview server and navigate to `step4_form.html`.
After user submits, read the result:

```javascript
// via preview_eval:
window.__result
// returns: { submitted, table_type, banner, stub_mode, stub }
```

`window.__result` structure:
```json
{
  "submitted": true,
  "table_type": 4,
  "banner": ["Q1", "Q2"],
  "stub_mode": "all",
  "stub": ["Q1", "Q2", "Q3", "..."]
}
```

Table type mapping: `1` = Count only · `2` = Pct only · `3` = Pct + Sig · `4` = All

#### Option B — Sequential text questions (fallback)

Ask 3 questions one by one when preview panel is not available:

**Q1 — Table type:**
> "Bạn muốn chạy bảng theo dạng nào?
> 1. Count only
> 2. Percentage only
> 3. Percentage + Sig test (90% & 95%)
> 4. Tất cả (Count + Percentage + Sig test)"

**Q2 — Banner:** Show SA/MA questions from metadata.json with their labels. Ask which ones.
→ Create `banner` with Total + chosen questions. Use `metadata.json` for choice codes.

**Q3 — Stub:** Show all codeable questions (SA, MA, Matrix) from metadata.json.
→ "all" → include all sorted by position; specific list → include in user's order.

Save to `output/SURVEY_NAME/datatable/datatable.json`.

### Step 5 — Run pipeline (table-only)

Since `data/` exists from Step 3, always run table-only:

```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    v1
```

**Force re-ingestion** (after fetching new MCP data):
```bash
python run_pipeline.py \
  --mcp-dir         output/SURVEY_NAME/mcp \
  --output-dir      output/SURVEY_NAME \
  --version         vX \
  --force-ingestion
```

---

## Workflow B — User requests changes

1. Read `output/SURVEY_NAME/datatable/datatable.json`
2. Modify per user request
3. Save
4. Run pipeline (increment version, table-only):
   ```bash
   python run_pipeline.py --output-dir output/SURVEY_NAME --version vX
   ```

---

## datatable.json structure

`datatable.json` is an **array** — each item produces its own set of sheets in the xlsx.

```json
[
  {
    "title": "SURVEY_NAME - Data Table",
    "sub_title": "General",
    "significance_test": { "enabled": true, "levels": [90, 95], "method": "independent" },
    "banner": [
      { "label": "Total", "filter": null },
      {
        "label": "Gender",
        "question": "Q10",
        "groups": [
          { "label": "Male",   "value": 2 },
          { "label": "Female", "value": 1 }
        ]
      }
    ],
    "stub": [
      { "question": "Q10", "label": "Gender",    "stats": ["base", "percent"] },
      { "question": "Q36", "label": "Food Freq", "stats": ["base", "percent", "t2b", "b2b", "mean"] }
    ],
    "tables": [
      { "sheet": "Count", "cell_content": "count",      "show_sig": false, "enabled": true },
      { "sheet": "Pct",   "cell_content": "percentage", "show_sig": false, "enabled": true },
      { "sheet": "Sig",   "cell_content": "percentage", "show_sig": true,  "enabled": true }
    ]
  }
]
```

Sheet tab name = `{sub_title} - {sheet}` → e.g. `"General - Count"`, `"General - Pct"`.

---

## Banner rules

- Always include `{ "label": "Total", "filter": null }` as first entry
- `value` = single integer code; `values` = list of codes (for grouping)
- Use the question **label** (e.g. `"Q10"`) as the `question` field
- To find choice codes → read `output/SURVEY_NAME/data/metadata.json` → find question → inspect `choices_i18n`
- MA questions supported as banner — each code becomes a column

### banner_matrix — Matrix rows as nested banner columns

Use `banner_matrix` to add a matrix question's rows (e.g. brands) as the innermost level of every banner column. When active, stub matrix questions automatically use **paired mode** — instead of expanding by sub-question, they show choice distributions matched to each brand column.

Without `groups` — all rows from `choices_i18n.rows` become individual columns:
```json
"banner_matrix": {
  "label": "Brand",
  "question": "Q17"
}
```

With explicit `groups`:
```json
"banner_matrix": {
  "label": "Brand",
  "question": "Q17",
  "groups": [
    { "label": "International brands", "row_codes": ["1","2","3","4","5","6"] },
    { "label": "Castrol",  "row_code": "1"  },
    { "label": "Shell",    "row_code": "2"  },
    { "label": "SK ZIC",   "row_code": "10" },
    { "label": "GS KIXX",  "row_code": "11" }
  ]
}
```

**Groups:**
- `row_code` (string): single brand → reads `{q_col}_r{code}` per column (paired mode)
- `row_codes` (list): grouped brands → aggregates across all `{q_col}_r{rc}` (stacked mode)
- Mix of both is allowed

**Example prompts that trigger banner_matrix:**
- "Brand làm header cuối, tất cả brands từ Q17"
- "Chạy bảng matrix: Castrol và Shell riêng lẻ, nhóm International brands gồm rows 1-6"
- "Dùng Q17 làm brand header column"

---

## Stub rules

- `stats` options: `"base"`, `"percent"`, `"t2b"`, `"b2b"`, `"mean"`, `"std"`, `"se"`
- Supported types: `SA`, `MA`, `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`
  - Matrix questions expand into one block per sub-question (row)
  - When `banner_matrix` is active → paired mode: show choices per brand column instead
- Excluded automatically: `FT`, `NUM`, `instruction`, `user-name`, `user-phone`, `reward`, `record`
- Output order follows the order in `stub`

### row_group — Matrix questions with shared row headers

Use `row_group: true` to group multiple matrix questions that share the **same** `choices_i18n.rows`. Renders one section header per row label, with sub-blocks per question underneath.

```json
{
  "row_group": true,
  "items": [
    { "question": "Q13_1", "label": "Motorbike oil — brands", "stats": ["base"] },
    { "question": "Q13_2", "label": "4-Wheel oil — brands",   "stats": ["base"] },
    { "question": "Q17",   "label": "Distributor satisfaction","stats": ["base", "percent"] }
  ]
}
```

**Rules:**
1. All items must be matrix questions (`Matrix_SA`, `Matrix_MA`, `Matrix_NUM`)
2. All items must share identical `choices_i18n.rows`
3. Cannot mix non-matrix questions inside a `row_group`

### Sub-question reference — Q{label}_r{n}

To add a single row from a matrix as a flat stub entry:

```json
{ "question": "Q17_r10",      "label": "SK ZIC — satisfaction",    "stats": ["base", "percent"] },
{ "question": "Q14_Freq_r10", "label": "SK ZIC — visit frequency", "stats": ["base", "percent"] }
```

- `r10` = sub-question with `row_index == 10` in metadata
- Works for any matrix type, outside `row_group`

---

## Tables rules

- `enabled: false` → sheet is skipped
- Choice mapping:
  - Count only → `[{ "sheet": "Count", "cell_content": "count", "show_sig": false }]`
  - Pct only   → `[{ "sheet": "Pct",   "cell_content": "percentage", "show_sig": false }]`
  - Sig        → add `{ "sheet": "Sig", "cell_content": "percentage", "show_sig": true }`
  - All        → Count + Pct + Sig

---

## Common requests → datatable.json changes

| User says | Claude does |
|---|---|
| "thêm income vào banner" | Add question to `banner` array |
| "bỏ Q15 khỏi stub" | Remove Q15 from `stub` |
| "thêm mean std cho Q36" | Add `"mean"`, `"std"` to Q36's `stats` |
| "thêm tất cả câu vào stub" | Add all codeable questions |
| "tắt sig test" | Set `significance_test.enabled = false` |
| "chỉ chạy 1 sheet percentage" | Modify `tables` to keep only Pct sheet |
| "nhóm Q13/Q14/Q17 theo brand" | Use `row_group: true` with those questions |
| "thêm SK ZIC riêng cho Q17" | Add `Q17_r{n}` sub-question ref to stub |
| "brand làm header cuối, tất cả brands" | Add `banner_matrix: { question: "QX" }` |
| "brand header, Castrol riêng, nhóm International" | Add `banner_matrix` with `groups` |
| "bảng bình thường + bảng matrix brand" | Two items in array: one without, one with `banner_matrix` |
| "refresh data / lấy data mới" | Re-fetch MCP → re-run with `--force-ingestion` |

---

## Confirm before acting

Before running pipeline or modifying datatable.json, always summarize and ask user to confirm:

> "Tôi sẽ [thêm Q5 vào banner / sửa datatable.json / chạy pipeline v3]. Bạn xác nhận không?"

- "yes / ok / xác nhận / làm đi / chạy đi" → proceed
- "no / thôi / đổi lại" → stop and ask what to change
