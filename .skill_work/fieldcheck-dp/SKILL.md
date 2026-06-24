---
name: fieldcheck-dp
description: >
  Workflow skill for processing survey data using the surveyflow Python package.
  Use this skill whenever the user mentions running a survey, processing QMe data,
  building a datatable, or working with surveyflow — even if they don't say "surveyflow"
  explicitly. Triggers on phrases like "chạy survey", "tạo datatable", "thêm banner",
  "bỏ câu khỏi stub", "fetch data từ QMe", "chạy pipeline", "chạy bảng", etc.
  Always use this skill for any QMe survey data task — it contains the authoritative
  workflow guide that governs all fetch, ingestion, and table steps.
---

# SurveyFlow Skill

This skill guides Claude through the full surveyflow pipeline:
**Fetch MCP → Ingestion → Design datatable.json → Table (xlsx)**

---

## How to start

User có thể bắt đầu bằng nhiều cách:

```
/fieldcheck-dp run pipeline VN8966
/fieldcheck-dp run pipeline VN8894 - Express
chạy survey VN8966
làm bảng cho VN8966
fetch data VN8894
chạy quality check VN8966
```

**SURVEY_NAME** = tên folder output, thường là:
- Survey code ngắn: `VN8966`
- Hoặc tên đầy đủ: `VN8894 - Express` (nếu tên có dấu cách, dùng quotes trong CLI)

Nếu user không nói tên survey → hỏi ngay: *"Bạn muốn chạy survey nào?"*

Nếu user nói **"hướng dẫn"**, **"help"**, **"dùng như thế nào"**, **"giải thích"** → đọc `skills/surveyflow/USER_GUIDE.md` và present toàn bộ nội dung cho user.

---

## Environment check (once per session)

```bash
pip install surveyflow mcp --upgrade --no-deps --break-system-packages -q
python -c "import surveyflow; print(surveyflow.__version__)"
```

For local dev: `pip install -e . --no-deps --break-system-packages -q`

---

## Output folder structure

```
output/SURVEY_NAME/
├── mcp/                    ← raw MCP files (fetch once, reuse)
│   ├── definition.json
│   └── data_export.csv
├── data/                   ← rawdata.csv + metadata.json (ingestion output, reused)
├── datatable/
│   └── datatable.json      ← Claude manages this
├── quality/                ← quality check output (optional)
│   ├── quality_report.json
│   └── flagged_profiles.csv
├── v1/
│   ├── datatable.xlsx      ← bảng crosstab
│   ├── chart_data.json     ← tự sinh kèm datatable.xlsx, dùng cho PPTX
│   └── slides.pptx         ← PPTX appendix (nếu đã tạo — Workflow D)
├── v2/
│   ├── datatable.xlsx
│   ├── chart_data.json
│   └── slides.pptx
└── ...
```

---

## Workflow A — First run

### Step 0 — Show progress tracker

Immediately show progress tracker, then execute steps one by one:

```
📋 Pipeline: SURVEY_NAME
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⏳ 1. Tìm survey         — đang tìm...
⬜ 2. Fetch data
⬜ 3. Ingestion
⬜ 4. Chọn cột / hàng
⬜ 5. Chạy bảng
⬜ 6. PPTX appendix      (tuỳ chọn)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

Icons: `⏳` đang chạy · `✅` xong · `⬜` chờ · `❌` lỗi

If quality check is requested, add it between steps 3 and 4:
```
✅ 3. Ingestion          — 450 responses
⏳ 3b. Quality check    — đang chạy...
⬜ 4. Chọn cột / hàng
```

Always end each update with a **clear next-action hint** for the user.

---

### Step 1 — Find survey

```
search_surveys(query="SURVEY_NAME")
```

Note the `survey_id`. If multiple results → show list and ask user to confirm which one.

---

### Step 2 — Fetch MCP data

**First**, check if `output/SURVEY_NAME/mcp/` already has `definition.json` + `data_export.csv`:
- **Files exist** → Ask: *"Data đã có sẵn. Dùng data cũ hay fetch lại?"*
- **Files missing** → Fetch immediately (no confirmation needed)

> ⚠️ **Fetch rules:**
> - **ALWAYS** `format="code"` — NEVER `format="text"`
> - Never use `get_survey_rows`
> - MCP returns structuredContent (dict). Use Write tool with `json.dumps(result, ensure_ascii=False, indent=2)`
> - `data_export.csv`: write with `encoding="utf-8-sig"` (BOM for Excel)
> - **Profile status**: default `approved` only. If user wants pending profiles: note `--profile-status approved,pending`

**Step A — Definition:**
```
get_survey_definition(survey_id)
  → save to output/SURVEY_NAME/mcp/definition.json
```

**Step B — Export CSV:**
```
# 1. Trigger export job
prepare_survey_data_file(survey_id, format="code", force_refresh=False)

# 2. Poll until ready (repeat every retry_after_seconds)
get_survey_data_file_status(job_id)

# If stuck (no change after 3+ polls) → STOP immediately → tell user:
# "Job bị stuck. Vui lòng export file zip từ Fieldcheck và upload vào đây."
# → Continue with Workflow A — Fallback (upload zip) below

# 3. Read all chunks
read_survey_data_file(job_id, file="data", offset=0,    limit=500)
read_survey_data_file(job_id, file="data", offset=500,  limit=500)
... keep reading until pagination.has_more == false

# 4. Assemble chunks → save as data_export.csv (encoding="utf-8-sig")
```

After fetch: tell user `"✅ Fetch xong — {N} responses"`.

---

### Step 2 fallback — Upload zip (khi fetch bị stuck)

User uploads a zip file exported manually from Fieldcheck.

```python
import zipfile, os

with zipfile.ZipFile('uploaded.zip') as z:
    data_file = next(
        (n for n in z.namelist() if n.startswith('code_retail_report_')),
        None
    )
    if not data_file:
        raise ValueError("Không tìm thấy file code_retail_report_* trong zip")
    with z.open(data_file) as src:
        content = src.read().decode('utf-8-sig')

os.makedirs('output/SURVEY_NAME/mcp', exist_ok=True)
with open('output/SURVEY_NAME/mcp/data_export.csv', 'w', encoding='utf-8-sig') as dst:
    dst.write(content)
```

> Nếu `definition.json` chưa có: fetch bằng `get_survey_definition(survey_id)` trước.

Continue to Step 3.

---

### Step 3 — Run ingestion

Run **immediately after `definition.json` is saved** — do not wait for `data_export.csv`.
This generates `metadata.json` so the datatable builder can be shown to user while CSV is still fetching.

**Full ingestion** (when `data_export.csv` is ready):
```bash
python run_pipeline.py \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --export-csv output/SURVEY_NAME/mcp/data_export.csv \
  --output-dir output/SURVEY_NAME
```

**Metadata-only** (when CSV not yet available — generates metadata.json from definition only):
```bash
python run_pipeline.py \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --output-dir output/SURVEY_NAME
```
> Full ingestion will run again once data_export.csv is ready.

**Output:**
- `data/rawdata.csv`
- `data/metadata.json` ← question labels + choice codes

After ingestion, always tell user:

```
✅ Ingestion xong — {N} rows, metadata.json sẵn sàng.

💡 Bạn có muốn chạy quality check trước khi tạo bảng không?
   Quality check sẽ kiểm tra toàn bộ {N} respondents xem có ai:
   - Bỏ qua câu hỏi bắt buộc
   - Được route đến câu nhưng không trả lời
   - Trả lời câu không được hiển thị
   - Có câu trả lời mâu thuẫn nhau

   Gõ "chạy quality" hoặc "bỏ qua" để tiếp tục tạo bảng.
```

**Chỉ chạy quality check khi user xác nhận** — không tự động chạy.

---

### Step 3b — Quality check (user confirms)

Chỉ chạy khi user đồng ý (sau khi được thông báo ở Step 3).
Cũng có thể chạy bất cứ lúc nào khi user nói: *"chạy quality", "kiểm tra data", "check lỗi routing"*.

```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --run-quality
```

#### After running — always present summary

Read `quality_report.json` and present:

```
📊 Quality Report — SURVEY_NAME
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Tổng respondents  : {total_respondents}
Bị flag           : {flagged_count} profile ({pct:.1f}%)

Loại vi phạm:
  ❌ missing          {n}  — câu luôn hiển thị nhưng không có trả lời
  ⚠️  routed_missing   {n}  — được route đến nhưng không trả lời
  🔍 extraneous       {n}  — trả lời câu không được hiển thị
  💥 contradiction    {n}  — câu trả lời mâu thuẫn nhau

Câu bị flag nhiều nhất:
  {Q_label} ({question_text})  — {N} lần  (missing: n, routed_missing: n)
  ...top 5...

Phạm vi kiểm tra:
  show_condition : {n_sc} câu
  contradiction  : {n_cs} câu
  always shown   : {n_always} câu
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📄 Chi tiết: output/SURVEY_NAME/quality/flagged_profiles.csv
```

**Summary rules:**
- `pct = flagged_count / total_respondents * 100`
- Top 5 questions: group by `question`, count by type, sort by total desc
- If `flagged_count == 0` → `"✅ Không có vi phạm nào."`

**Actionable next steps — always show after summary:**

| Loại vi phạm chủ yếu | Gợi ý |
|---|---|
| `missing` optional | Bình thường — không cần xử lý |
| `missing` mandatory | Báo cáo team — profile có thể cần loại |
| `routed_missing` nhiều | Khả năng lỗi survey logic — check với PM |
| `extraneous` | Lỗi routing QMe — nên báo cáo |
| `contradiction` | Profile vi phạm logic — thường loại trước khi chạy bảng |

Sau đó hỏi: *"Bạn muốn xem chi tiết câu nào, hay tiếp tục chạy bảng?"*

#### Drill-down khi user hỏi thêm

**"Xem chi tiết câu Q5"** → filter `violations` where `question == "Q5"`, show table (10 rows):
```
profile_id | type           | detail                          | condition_trigger
-----------|----------------|---------------------------------|------------------
123456     | routed_missing | condition met but no answer...  | Q3 in [1,2]
```

**"Profile nào bị lỗi nhiều nhất?"** → group by `profile_id`, count, show top 10.

**"Chỉ xem contradiction"** → filter by `type == "contradiction"`.

---

### Step 4 — Design datatable.json

**If user already specified requirements** → create `output/SURVEY_NAME/datatable/datatable.json` directly, then go to Step 5.

**If not specified** → render the **Datatable Builder artifact**:

```python
# Read the HTML builder from skill assets
with open("skills/surveyflow/datatable_builder.html", encoding="utf-8") as f:
    html_content = f.read()
# Render as artifact in chat
```

The builder is a **multi-step web UI** that runs in the artifact panel:

```
Step 1: Cài đặt  → Tiêu đề bảng + Loại output (Pct/Count/Sig)
Step 2: Banner   → Chọn câu SA/MA làm cột header
Step 3: Stub     → Chọn câu SA/MA/Matrix làm hàng
Step 4: Xác nhận → Review + nút "Tạo datatable.json"
```

**User flow:**
1. User mở builder → kéo/thả `output/SURVEY_NAME/data/metadata.json` vào ô upload
2. Chọn cài đặt → banner → stub → xác nhận
3. Nhấn **"Tạo datatable.json"** → JSON hiện ra trong trang với nút Sao chép
4. User sao chép và paste vào chat — Claude nhận được message dạng:

```
[DATATABLE_CONFIG]
{ ... json ... }
```

**Khi Claude nhận `[DATATABLE_CONFIG]`:**
1. Parse JSON từ message
2. Save to `output/SURVEY_NAME/datatable/datatable.json`
3. Confirm with user: *"Đã lưu datatable.json. Chạy v1 nhé?"*
4. Run pipeline (Step 5)

> ⚠️ Builder chỉ tạo cấu hình cơ bản (banner groups, stub list, output type).
> Các tính năng nâng cao (show_total, banner_matrix, row_group, custom_ref, sub-question ref)
> được thêm sau thông qua Workflow B.

**Helper scripts** (Claude dùng internally để tra cứu choice codes khi cần):
```bash
python skills/surveyflow/scripts/list_questions.py output/SURVEY_NAME/data/metadata.json --type banner
python skills/surveyflow/scripts/list_questions.py output/SURVEY_NAME/data/metadata.json --type stub
python skills/surveyflow/scripts/check_choices.py  output/SURVEY_NAME/data/metadata.json --question Q10
```

---

### Step 5 — Run pipeline

```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    v1
```

With language:
```bash
python run_pipeline.py --output-dir output/SURVEY_NAME --version v1 --lang en   # English
python run_pipeline.py --output-dir output/SURVEY_NAME --version v1 --lang vi   # Vietnamese (default)
```

Force re-ingestion after new fetch:
```bash
python run_pipeline.py \
  --mcp-dir         output/SURVEY_NAME/mcp \
  --export-csv      output/SURVEY_NAME/mcp/data_export.csv \
  --output-dir      output/SURVEY_NAME \
  --version         vX \
  --force-ingestion
```

> ⚠️ NEVER recreate or rewrite `run_pipeline.py` — it is always available in the project root.

#### After pipeline — always present result

```
✅ Datatable xong!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 File   : output/SURVEY_NAME/v1/datatable.xlsx
📊 Sheets : General - Pct  |  General - Count  |  General - Sig
👥 Rows   : 450 respondents
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bạn muốn thay đổi gì không? (thêm banner, thêm câu, đổi ngôn ngữ, tạo PPTX appendix...)
```

---

### Step 6 — PPTX Chart Appendix (tuỳ chọn)

Sau khi pipeline chạy xong, `chart_data.json` đã có sẵn cạnh `datatable.xlsx`.
Khi user nói *"tạo appendix"*, *"chạy slides"*, *"tạo PPTX"* → xác nhận rồi chạy.

**Cách 1 — kèm pipeline** (table + appendix trong 1 lệnh):
```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    vX \
  --appendix
```

**Cách 2 — riêng** (sau khi đã có chart_data.json):
```bash
surveyflow-pptx \
  output/SURVEY_NAME/vX/chart_data.json \
  output/SURVEY_NAME/vX/slides.pptx
```

Tuỳ chọn `surveyflow-pptx`:
- `--table N` — chỉ chạy bảng thứ N (0-indexed)
- `--start-page N` — bắt đầu đánh số trang từ N

> ⚠️ Self-contained — không cần `documents/temp.pptx`. Style từ `surveyflow/steps/appendix/chart_templates/`.  
> ⚠️ **NEVER recreate or rewrite `surveyflow/steps/appendix/generate_pptx.py`** — nằm sẵn trong package.

After running:
```
✅ PPTX xong!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 File : output/SURVEY_NAME/vX/slides.pptx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## Workflow D — PPTX Chart Appendix (standalone)

Khi user muốn tạo PPTX từ kết quả đã có (không phải trong Workflow A):

**Confirm trước khi chạy:**
> *"Tôi sẽ tạo PPTX appendix từ `output/SURVEY_NAME/vX/chart_data.json`. Bạn xác nhận không?"*

**Cách 1 — chạy lại pipeline kèm appendix:**
```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    vX \
  --appendix
```

**Cách 2 — chạy riêng từ chart_data.json đã có:**
```bash
surveyflow-pptx \
  output/SURVEY_NAME/vX/chart_data.json \
  output/SURVEY_NAME/vX/slides.pptx
```

**Chart type tự chọn:**
- `donut_stacked` → donut Total + 100%-stacked breakdown (SA ≤ 5 choices)
- `bar_vertical` → cột dọc Total + cột breakdown (SA > 5 choices)
- `bar_horizontal` → bar ngang Total + cột breakdown (MA)

> ⚠️ Self-contained — không cần `documents/temp.pptx`.  
> Style: `surveyflow/steps/appendix/chart_templates/{bar,col,donut,stacked}.xml`  
> ⚠️ **NEVER recreate or rewrite `surveyflow/steps/appendix/generate_pptx.py` hoặc `tools/extract_chart_templates.py`**  
> CLI command sau `pip install surveyflow`: `surveyflow-pptx`

---

## Workflow B — User requests changes

When user asks to modify the table:

1. **Read** current `output/SURVEY_NAME/datatable/datatable.json`
2. **Understand** the request (see Common requests table below)
3. **Confirm** change before modifying: *"Tôi sẽ [mô tả thay đổi]. OK không?"*
4. **Modify and save** `datatable.json`
5. **Detect current version**: check which `vX/` folders exist → use next number
6. **Run pipeline** with new version (table-only, no `--mcp-dir` needed)
7. **Present result** using same template as Step 5

**After each change:**
> *"Đã cập nhật v2. File cũ v1 vẫn còn. Bạn muốn thay đổi thêm gì không?"*

**Version rules:**
- v1 = first table
- Increment (v1→v2→v3...) every time datatable.json changes and pipeline re-runs
- NEVER overwrite existing vX folder — always create new one
- Always tell user: *"Tôi sẽ tạo v2 — file v1 vẫn giữ nguyên"*

---

## datatable.json structure

`datatable.json` is an **array** — each item produces its own set of sheets.

**Item types:**
- `{ "type": "datatable", ... }` → cross-tab table (or omit `type`, default = datatable)
- `{ "_custom_defs": [...] }` → reusable filter definitions (no sheet output)

```json
[
  {
    "type": "datatable",
    "title": "SURVEY_NAME - Data Table",
    "sub_title": "General",
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
      { "sheet": "Pct",   "cell_content": "percentage", "show_sig": false, "enabled": true, "decimal": 0 },
      { "sheet": "Sig",   "cell_content": "percentage", "show_sig": true,  "levels": [90, 95], "method": "independent", "enabled": true }
    ]
  }
]
```

Sheet tab = `{sub_title} - {sheet}` → e.g. `"General - Count"`.

---

## Banner rules

- Always include `{ "label": "Total", "filter": null }` as first entry
- `value` = single integer code; `values` = list of codes (for grouping)
- Use the question **label** (e.g. `"Q10"`) as the `question` field — from `metadata.json`
- MA questions supported as banner
- To find choice codes → read `metadata.json` → find question → `choices_i18n`

### show_total — add Total before sub-groups

```json
{
  "label": "Area", "question": "S3a",
  "groups": [{ "label": "HN", "value": 1 }, { "label": "HCM", "value": 2 }],
  "show_total": true
}
```

If **any** entry has `show_total: true`, the global Total column is hidden.

### _custom_defs + custom_ref — reusable filter groups

Define once:
```json
{
  "_custom_defs": [
    {
      "label": "Users",
      "choices": [
        { "code": 1, "label": "Brand A", "filter": { "question": "S11b", "codes": [1, 2] } },
        { "code": 2, "label": "Brand B", "filter": { "and": [
            { "question": "S11b", "codes": [3, 4] },
            { "question": "S4",   "codes": [2, 3] }
        ]}}
      ]
    }
  ]
}
```

Reference in `banner`: `{ "type": "custom_ref", "ref": "Users", "label": "Users" }`

Nested (Gender × Users):
```json
{
  "label": "Gender", "question": "S5",
  "groups": [{ "label": "Nam", "value": 1 }, { "label": "Nữ", "value": 2 }],
  "levels": [{ "type": "custom_ref", "ref": "Users", "label": "Users" }]
}
```

**Filter syntax:** `{ "question": "Qx", "codes": [...] }` / `{ "and": [...] }` / `{ "or": [...] }`

### custom_ref in stub

```json
{ "type": "custom_ref", "ref": "Users", "stats": ["base", "percent"] }
```

### banner_matrix — Matrix rows as nested columns

```json
"banner_matrix": { "label": "Brand", "question": "Q17" }
```

With groups (`row_code`/`row_codes` accept integers or strings):
```json
"banner_matrix": {
  "label": "Brand", "question": "Q17",
  "groups": [
    { "label": "International", "row_codes": [1, 2, 3] },
    { "label": "Castrol",       "row_code":  1 },
    { "label": "Shell",         "row_code":  2 }
  ]
}
```

### matrix_orientation — Matrix rows as banner sub-columns

Khi toàn bộ stub là matrix questions có cùng rows:

```json
{
  "matrix_orientation": "horizontal",
  "stub": [
    { "question": "Q17", "label": "Brand Satisfaction", "stats": ["base", "percent"] },
    { "question": "Q19", "label": "Brand Imagery",      "stats": ["base", "percent"] }
  ]
}
```

Rows → banner sub-columns (Total/Brand A, Total/Brand B...), Choices → stub rows.

### matrix_rows — show/hide/combine brand sub-columns

Dùng cùng `matrix_orientation: "horizontal"` để chọn brands hiện:

```json
"matrix_rows": [
  { "row_code": 1,             "label": "STING" },
  { "row_code": 2,             "label": "RED BULL" },
  { "row_codes": [4,5,6,7],   "label": "Others" }
]
```

Rows không liệt kê → tự động ẩn. Bỏ qua `matrix_rows` → hiện tất cả.

### sig_direction

```json
"sig_direction": "rows"
```

- `"rows"` (default): so sánh brands với nhau trong cùng demographic
- `"columns"`: so sánh demographics với nhau trong cùng brand

---

## Stub rules

- `stats`: `"base"`, `"percent"`, `"t2b"`, `"b2b"`, `"mean"`, `"std"`, `"se"`
- Types: `SA`, `MA`, `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`
- Auto-excluded: `FT`, `NUM`, `instruction`, `user-name`, `user-phone`, `reward`, `record`
- Output order follows `stub` array order

### row_group — Matrix questions with shared row headers

```json
{
  "row_group": true,
  "items": [
    { "question": "Q13_1", "label": "Motorbike oil", "stats": ["base"] },
    { "question": "Q17",   "label": "Satisfaction",  "stats": ["base", "percent"] }
  ]
}
```

All items must be matrix questions with identical `choices_i18n.rows`.

### Sub-question reference

```json
{ "question": "Q17_r10", "label": "SK ZIC — satisfaction", "stats": ["base", "percent"] }
```

---

## Tables rules

- `enabled: false` → sheet is skipped
- `decimal`: `0` → "0%" (default), `1` → "0.0%", `2` → "0.00%"
- Sig test config **per sheet** (top-level `significance_test` block is ignored):
  - `show_sig: true` + `levels` (default `[90, 95]`) + `method` (default `"independent"`)

**Mapping:**
```
Count only → [{ sheet: Count, count,      no_sig }]
Pct only   → [{ sheet: Pct,   percentage, no_sig, decimal: 0 }]
Sig        → [{ sheet: Sig,   percentage, show_sig: true, levels:[90,95], method: independent }]
All        → Count + Pct + Sig
```

---

## Language selection (`--lang`)

Controls question and choice labels in output xlsx. Pass at runtime (not stored in datatable.json):

```bash
python run_pipeline.py --output-dir output/VN8966 --version v1 --lang en   # English
python run_pipeline.py --output-dir output/VN8966 --version v1 --lang vi   # Vietnamese (default)
```

> 💡 Nếu user muốn output tiếng Anh → nhớ thêm `--lang en` khi chạy Step 5 hoặc Workflow B.

---

## Common requests → datatable.json changes

| User says | Claude does |
|---|---|
| "thêm income vào banner" | Add question to `banner` array |
| "bỏ Q15 khỏi stub" | Remove Q15 from `stub` |
| "thêm mean std cho Q36" | Add `"mean"`, `"std"` to Q36's `stats` |
| "thêm tất cả câu vào stub" | Add all codeable questions |
| "tắt sig test" | Set `"show_sig": false` on Sig sheet |
| "bật sig test 90% và 95%" | Add `"show_sig": true, "levels": [90, 95]` to Sig sheet |
| "hiện 1 chữ số thập phân" | Add `"decimal": 1` to Pct sheet |
| "thêm total cho từng banner group" | Add `"show_total": true` to banner entry |
| "chỉ chạy 1 sheet percentage" | Keep only Pct sheet in `tables` |
| "nhóm Q13/Q14/Q17 theo brand" | Use `row_group: true` |
| "thêm SK ZIC riêng cho Q17" | Add `Q17_r{n}` sub-question ref |
| "tạo filter group dùng chung" | Create `_custom_defs` item |
| "thêm user groups vào banner" | Add `{ "type": "custom_ref", "ref": "Name" }` |
| "user groups × area nested" | Use `levels: [{ "type": "custom_ref", ... }]` |
| "brand làm header, tất cả brands" | Add `banner_matrix: { question: "QX" }` |
| "brand header, Castrol riêng, nhóm International" | Add `banner_matrix` with `groups` |
| "2 bảng: bình thường + matrix brand" | Two items in array |
| "chạy bảng matrix horizontal" | Add `"matrix_orientation": "horizontal"`; stub chỉ matrix questions cùng rows |
| "chỉ hiện 4 brands, nhóm others" | Add `matrix_rows` với `row_code` riêng + `row_codes` nhóm |
| "ẩn brand X" | Xoá row_code đó khỏi `matrix_rows` |
| "refresh data / lấy data mới" | Re-fetch → re-run with `--force-ingestion` |
| "tạo appendix PPTX / chạy slides" | Workflow D: `surveyflow-pptx output/SURVEY_NAME/vX/chart_data.json output/SURVEY_NAME/vX/slides.pptx` |
| "chạy quality check" | Run `--run-quality`, present summary |
| "xem chi tiết câu Q5" | Filter violations by question, show table |
| "profile nào bị lỗi nhiều nhất" | Group by profile_id, show top 10 |
| "export tiếng Anh" | Run with `--lang en` |
| "include pending profiles" | Add `--profile-status approved,pending` |

---

## FT Coding Workflow

Khi user nói: *"code câu FT"*, *"code Q26"*, *"code open-ended"* → chạy workflow này. Không dùng artifact — toàn bộ bằng bash_tool.

> ⚡ **Effort requirement:** Trước khi bắt đầu, **luôn nhắc user chuyển Effort lên Medium trở lên** (Model picker → Effort → Medium/High/Max). FT coding đòi hỏi Claude đọc hàng trăm responses, tạo code frame MECE và viết regex chính xác — Effort thấp (Low/Default) sẽ cho kết quả kém. Chỉ tiếp tục sau khi user xác nhận đã chuyển.
>
> *"Trước khi code, bạn vui lòng chuyển **Effort lên Medium trở lên** (Model picker → Effort → Medium). Đã chuyển chưa?"*

### Overview

FT (free-text) questions không thể đưa vào datatable trực tiếp. Pipeline gồm 7 bước:
1. List câu FT, hỏi user chọn câu nào
2. Đọc valid responses từ rawdata
3. Tạo code frame + assign codes bằng regex (rule-based, không cần API key)
4. Add binary columns vào rawdata
5. Inject câu MA vào metadata.json
6. Thêm stub vào datatable.json
7. Chạy bảng (table-only — KHÔNG force-ingestion)

---

### Step FT-1 — List câu FT, hỏi user

```python
import json

with open('output/SURVEY_NAME/data/metadata.json') as f:
    meta = json.load(f)

ft_qs = sorted(
    [(q['label'], q.get('question_i18n', {}).get('vi', ''))
     for q in meta['questions'].values()
     if q.get('answer_type') == 'FT'],
    key=lambda x: x[0]
)
for label, vi in ft_qs:
    print(f"  {label:10} {vi[:60]}")
```

Hiển thị danh sách, hỏi: *"Bạn muốn code câu nào?"*

---

### Step FT-2 — Đọc valid responses

```python
import json, pandas as pd

Q_LABEL = 'Q26'       # câu cần code — thay theo yêu cầu
SURVEY  = 'VN8963'    # thay theo survey

df = pd.read_csv(f'output/{SURVEY}/data/rawdata.csv', low_memory=False)

responses_series = df[Q_LABEL].dropna()
responses_series = responses_series[responses_series.str.strip().str.len() > 2]
responses_series = responses_series[responses_series.str.lower() != 'test']
responses = list(responses_series)

# Save for reuse in next steps
with open(f'/tmp/{Q_LABEL}_responses.json', 'w', encoding='utf-8') as f:
    json.dump(responses, f, ensure_ascii=False, indent=2)

print(f"{Q_LABEL}: {len(responses)} valid responses")
# Print first 30 to read before building code frame
for i, r in enumerate(responses[:30], 1):
    print(f"  {i:3}. {r[:100]}")
```

Claude đọc output rồi tự tạo CODE_FRAME và RULES ở bước tiếp theo.

---

### Step FT-3 — Tạo code frame + assign codes + summary

Claude đọc responses từ output FT-2, tự tạo CODE_FRAME (8–15 mã MECE) và RULES (regex tiếng Việt). Chạy toàn bộ trong 1 bash block:

```python
import re, json, pandas as pd

Q_LABEL = 'Q26'
SURVEY  = 'VN8963'

with open(f'/tmp/{Q_LABEL}_responses.json', encoding='utf-8') as f:
    responses = json.load(f)

# ── CODE FRAME — Claude tự điền ──────────────────────────────
CODE_FRAME = [
    (1,  "Label EN",       "Mô tả tiếng Việt"),
    (2,  "...",            "..."),
    # thêm codes tuỳ nội dung câu hỏi
    (99, "Other/Unclear",  "Khác/không rõ"),
]

# ── RULES — regex tiếng Việt, 1 pattern per code ─────────────
RULES = [
    (1,  r'pattern_a|pattern_b'),
    (2,  r'pattern_c|pattern_d'),
    # ...
]

def assign_codes(text):
    t = text.lower()
    assigned = [cid for cid, pattern in RULES if re.search(pattern, t)]
    return assigned if assigned else [99]

all_codes = [assign_codes(r) for r in responses]

# ── Summary ───────────────────────────────────────────────────
n = len(responses)
counts = {}
for codes in all_codes:
    for c in codes:
        counts[c] = counts.get(c, 0) + 1

print(f"\n📊 FT Coding — {Q_LABEL} (n={n})")
print("=" * 50)
for cid, label_en, label_vi in sorted(CODE_FRAME, key=lambda x: -counts.get(x[0], 0)):
    cnt = counts.get(cid, 0)
    print(f"  {cid:2}. {label_en:<28} {cnt:3} ({cnt/n*100:4.1f}%)")
```

---

### Step FT-4 — Add vào rawdata

Chạy tiếp trong cùng bash block (hoặc block mới — `all_codes` và `CODE_FRAME` phải còn trong scope):

```python
df = pd.read_csv(f'output/{SURVEY}/data/rawdata.csv', low_memory=False)

# Dùng pd.concat để tránh fragmentation warning
new_cols = {f'{Q_LABEL}_coded_{cid}': pd.Series(0, index=df.index)
            for cid, _, _ in CODE_FRAME}
df = pd.concat([df, pd.DataFrame(new_cols)], axis=1)

# Fill coded rows — map theo index gốc trong rawdata
q_mask = (df[Q_LABEL].notna()
          & (df[Q_LABEL].str.strip().str.len() > 2)
          & (df[Q_LABEL].str.lower() != 'test'))
q_indices = df[q_mask].index.tolist()

for i, idx in enumerate(q_indices):
    code_set = set(all_codes[i])
    for cid, _, _ in CODE_FRAME:
        df.at[idx, f'{Q_LABEL}_coded_{cid}'] = 1 if cid in code_set else 0

df.to_csv(f'output/{SURVEY}/data/rawdata.csv', index=False)
print(f"✅ rawdata: {len(df)} rows, {len(df.columns)} cols")
print(f"   New cols: {[f'{Q_LABEL}_coded_{c[0]}' for c in CODE_FRAME]}")
```

> ⚠️ **KHÔNG chạy `--force-ingestion` sau bước này** — sẽ overwrite rawdata từ export CSV gốc và mất toàn bộ cột coded.

---

### Step FT-5 — Inject câu MA vào metadata.json

```python
with open(f'output/{SURVEY}/data/metadata.json') as f:
    meta = json.load(f)

qs = meta['questions']

# Tìm position của câu FT gốc
ft_pos = next((q['position'] for q in qs.values()
               if q.get('label') == Q_LABEL), 999)

# Key dict phải là string duy nhất không trùng qid thật
fake_qid = f'coded_{Q_LABEL}'

qs[fake_qid] = {
    "position":    ft_pos + 0.5,
    "question_id": 900000 + hash(Q_LABEL) % 10000,
    "label":       f"{Q_LABEL}_coded",
    "question_i18n": {
        "vi": f"{Q_LABEL} - [Nội dung câu] (coded)",
        "en": f"{Q_LABEL} - [Question text] (coded)"
    },
    "answer_type": "MA",
    "mandatory":   False,
    "status":      1,
    "choices_i18n": {
        str(cid): {"vi": label_vi, "en": label_en}
        for cid, label_en, label_vi in CODE_FRAME
    },
    # CRITICAL: phải khớp chính xác tên cột trong rawdata
    "rawdata_columns": [f"{Q_LABEL}_coded_{cid}" for cid, _, _ in CODE_FRAME]
}

meta['questions'] = qs
with open(f'output/{SURVEY}/data/metadata.json', 'w', encoding='utf-8') as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)

print(f"✅ metadata: {len(qs)} questions")
print(f"   rawdata_columns: {qs[fake_qid]['rawdata_columns']}")
```

> **Naming rules — phải nhất quán:**
> - Rawdata columns: `{Q_LABEL}_coded_{choice_id}` — e.g. `Q26_coded_1`, `Q26_coded_99`
> - `rawdata_columns` trong metadata phải match chính xác tên cột trong rawdata
> - surveyflow đọc data qua `rawdata_columns` — sai tên → bảng có header nhưng không có số

---

### Step FT-6 — Thêm stub vào datatable.json

```python
with open(f'output/{SURVEY}/datatable/datatable.json') as f:
    dt = json.load(f)

# Xóa stub cũ nếu đã có (tránh duplicate khi re-run)
dt[0]['stub'] = [s for s in dt[0]['stub']
                 if s.get('question') != f'{Q_LABEL}_coded']

dt[0]['stub'].append({
    "question": f"{Q_LABEL}_coded",
    "label":    f"{Q_LABEL} - [Question text] (coded)",
    "stats":    ["base", "percent"]
})

with open(f'output/{SURVEY}/datatable/datatable.json', 'w', encoding='utf-8') as f:
    json.dump(dt, f, ensure_ascii=False, indent=2)

print(f"✅ datatable.json: {len(dt[0]['stub'])} stubs")
```

---

### Step FT-7 — Chạy bảng (table-only)

```python
from surveyflow.cli import main
import os, re as _re

# Detect next version
existing = [d for d in os.listdir(f'output/{SURVEY}') if _re.match(r'^v\d+$', d)]
next_v = f"v{max([int(d[1:]) for d in existing], default=0) + 1}"

main([
    '--output-dir',     f'output/{SURVEY}',
    '--profile-status', 'approved,pending',
    '--version',        next_v,
])
# KHÔNG dùng --mcp-dir, --export-csv, --force-ingestion
```

Present file sau khi xong bằng `present_files`.

---

### Fix percentage formatting (nếu cần)

Câu MA được inject vào metadata đôi khi render dưới dạng decimal (0.082) thay vì % (8%) trong sheet Pct. Kiểm tra và fix nếu cần:

```python
import openpyxl

xlsx_path = f'output/{SURVEY}/{next_v}/datatable.xlsx'
wb = openpyxl.load_workbook(xlsx_path)

for sheet_name in wb.sheetnames:
    if 'Count' in sheet_name:
        continue
    ws = wb[sheet_name]
    in_coded = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == f'{Q_LABEL}_CODED'.upper():
                in_coded = True
            if in_coded and isinstance(cell.value, float) and 0 < cell.value <= 1.0:
                cell.value = round(cell.value * 100, 1)
                cell.number_format = '0'

wb.save(xlsx_path)
print(f"✅ Percentage fix applied")
```

Chỉ chạy nếu mở file thấy số dạng `0.082` thay vì `8`.

---

### FT coding summary format

Luôn in sau FT-3 và trước FT-4:

```
📊 FT Coding — Q26 (n=526)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   6. Taste Improvement        110  (20.9%)
   2. Packaging Redesign       102  (19.4%)
   1. More Flavor Variety       85  (16.2%)
   3. More Advertising          70  (13.3%)
  99. Other/Unclear            103  (19.6%)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

Nếu Other/Unclear > 15% → đề xuất: *"Other/Unclear còn cao ({n}%). Bạn muốn tôi review và thêm rules không?"*

---

### Common requests → FT coding actions

| User says | Claude does |
|---|---|
| "code câu FT" | Chạy FT-1, list câu FT, hỏi user chọn |
| "code Q26" | Chạy FT-2 → FT-7 tự động cho Q26 |
| "code Q33c Q33d" | Chạy FT-2 → FT-7 lần lượt cho từng câu |
| "xem code frame" | In CODE_FRAME dạng bảng |
| "sửa code frame" | Update CODE_FRAME + RULES, re-run FT-4 → FT-7 |
| "review Other/Unclear" | Lấy responses có code 99, in 20 examples, đề xuất rules mới |
| "add câu FT đã code vào bảng" | Chỉ chạy FT-5 → FT-7 (skip FT-2/3/4 nếu rawdata đã có cols) |

---

## Error handling
, always:
1. Đọc traceback, xác định dòng lỗi
2. Giải thích ngắn gọn bằng tiếng Việt
3. Đề xuất fix cụ thể — không chỉ báo lỗi rồi dừng

| Error | Likely cause | Fix |
|---|---|---|
| `FileNotFoundError: rawdata.csv` | Ingestion chưa chạy | Chạy lại Step 3 |
| `KeyError: 'Q5'` trong pipeline | Question label sai trong datatable.json | Kiểm tra `metadata.json` để lấy label đúng |
| Output có 0 rows | `profile_status` filter quá hẹp | Thêm `--profile-status approved,pending` |
| `ModuleNotFoundError: surveyflow` | Package chưa install | `pip install surveyflow` |
| `PermissionError: datatable.xlsx` | File đang mở trong Excel | Đóng Excel trước khi chạy |
| `ValueError: No data_export.csv` | Fetch chưa hoàn thành | Fetch lại hoặc dùng upload zip |
| `json.JSONDecodeError` | datatable.json bị lỗi syntax | Đọc và kiểm tra file JSON |

---

## Confirm before acting

| Situation | Action |
|---|---|
| User gọi `/fieldcheck-dp run pipeline` | Chạy thẳng — không confirm |
| Fetch + ingestion | Chạy thẳng — không confirm |
| User nhấn nút trong artifact builder | Chạy pipeline ngay — không confirm |
| User yêu cầu sửa `datatable.json` | Confirm 1 lần: *"Tôi sẽ [thay đổi]. OK không?"* |
| Fetch lại data mới (overwrite file cũ) | Confirm — destructive action |
| Increment version | Thông báo: *"Tôi sẽ tạo v2 — v1 vẫn giữ nguyên"* |

- `"yes / ok / xác nhận / làm đi / chạy đi"` → proceed
- `"no / thôi / đổi lại"` → stop and ask what to change instead
