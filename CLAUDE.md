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

## ⚠️ Confirm before acting — ALWAYS

Before executing **any** of the following actions, summarize what Claude is about to do and ask user to confirm:

| Action | Confirm prompt example |
|---|---|
| Run pipeline | "Tôi sẽ chạy pipeline vX với config hiện tại. Bạn xác nhận chạy không?" |
| Modify `datatable.json` | "Tôi sẽ sửa datatable.json: [mô tả thay đổi]. Bạn xác nhận không?" |
| Fetch data from QMe | "Tôi sẽ fetch lại data từ QMe và ghi đè input files. Bạn xác nhận không?" |

**Rules:**
- Always confirm **before** acting, never after
- List all changes clearly so user knows exactly what will happen
- If user says "yes / ok / xác nhận / làm đi / chạy đi" → proceed
- If user says "no / thôi / đổi lại" → stop and ask what to change instead
- Exception: read-only actions (reading files, searching, checking version) do NOT need confirmation

> Note: Push GitHub / publish PyPI là việc của developer — không thuộc phạm vi hướng dẫn này.

---

## Environment setup (check once at start of every session)

```bash
python -c "import surveyflow; print(surveyflow.__version__)"
```

If fails → `pip install surveyflow` (or `pip install -e .` for local dev).

---

## Workflow A — First time running a survey

### Step 1 — Find survey
```
search_surveys(query="SURVEY_NAME")
```
Note the `survey_id`.

### Step 2 — Fetch and save MCP data

**First, check if `output/SURVEY_NAME/mcp/` already exists and contains `definition.json` + `data_export.csv`.**

If files exist → ask user:
> "Input data đã có sẵn (`definition.json`, `data_export.csv`). Bạn muốn dùng data cũ hay fetch lại data mới từ QMe?"

- User says **dùng data cũ / no** → skip to Step 3
- User says **fetch lại / yes** → proceed to fetch below

If files do not exist → fetch immediately (no need to ask):

**Step A — Definition:**
```
get_survey_definition(survey_id)
  → save to output/SURVEY_NAME/mcp/definition.json
```

**Step B — Export CSV (prepare → poll → read chunks):**
```
# 1. Trigger export job
prepare_survey_data_file(survey_id, format="code", force_refresh=False)
  → returns { job_id, status, expires_at, files[], ... }

# 2. Poll until ready (nếu status chưa phải "ready")
get_survey_data_file_status(job_id)
  → lặp lại mỗi retry_after_seconds cho đến khi status == "ready"

# 3. Read all chunks
read_survey_data_file(job_id, file="data", offset=0,    limit=500)
read_survey_data_file(job_id, file="data", offset=500,  limit=500)
... keep reading (offset += 500) until pagination.has_more == false

# 4. Assemble chunk text → save to output/SURVEY_NAME/mcp/data_export.csv
#    Write with encoding="utf-8-sig" (BOM for Excel)
```

> ⚠️ **MCP trả về structuredContent (dict), không phải plain text.**
> Dùng Write tool để ghi file — nội dung là `json.dumps(result, ensure_ascii=False, indent=2)`.
> Không dùng bash pipe cho MCP output.

> ⚠️ **Path nhất quán:** Trước khi lưu file đầu tiên, xác định base directory bằng:
> ```python
> import os; BASE = os.getcwd()
> ```
> Dùng `BASE` làm gốc cho **tất cả** các path trong cùng session.

### Step 3 — Run ingestion (generate rawdata + metadata)

**Check if `output/SURVEY_NAME/data/rawdata.csv` already exists.**

If exists → skip this step (data already generated).

If not exists → run ingestion. Since `datatable/datatable.json` doesn't exist yet,
the pipeline automatically skips the table step and only generates `data/`:

```bash
python run_pipeline.py \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --export-csv output/SURVEY_NAME/mcp/data_export.csv \
  --output-dir output/SURVEY_NAME
```

This produces:
- `output/SURVEY_NAME/data/rawdata.csv`
- `output/SURVEY_NAME/data/metadata.json`

> **Why separate ingestion first?**
> `metadata.json` contains the actual question labels, choice codes, and matrix row/column
> definitions needed to build a correct `datatable.json`. Ask datatable questions AFTER
> ingestion so you can reference real choice codes from metadata.

### Step 4 — Create datatable.json

**Nếu user đã chỉ định rõ yêu cầu** → tạo `datatable/datatable.json` theo yêu cầu đó.

**Nếu user chưa chỉ định** → KHÔNG auto-generate. Hỏi lần lượt 3 câu sau (hỏi từng câu, chờ user trả lời rồi mới hỏi câu tiếp):

**Câu hỏi 1 — Loại bảng:**
> "Bạn muốn chạy bảng theo dạng nào?
> 1. Count only
> 2. Percentage only
> 3. Percentage + Sig test (90% & 95%)
> 4. Tất cả (Count + Percentage + Sig test)"

→ Tạo `tables` trong datatable.json theo đúng lựa chọn. Không thêm sheet nào ngoài những gì user chọn.

**Câu hỏi 2 — Banner (header):**
> Hiển thị danh sách các câu SA/MA từ metadata.json, ví dụ:
> "Banner gồm những câu nào? (mặc định luôn có Total)
> Ví dụ: Q1 (Age), Q2 (Income), Q3 (Baby Age)
> Nhập số câu hoặc tên, cách nhau bằng dấu phẩy:"

→ Tạo `banner` với Total + các câu user chọn. Lấy choice codes từ `metadata.json`.

**Câu hỏi 3 — Stub:**
> Hiển thị danh sách tất cả câu codeable (SA, MA, Matrix) từ metadata.json:
> "Stub gồm những câu nào?
> - Nhập 'all' để lấy tất cả
> - Hoặc nhập số câu cách nhau bằng dấu phẩy: Q1, Q5, Q8, ..."

→ Nếu "all": thêm tất cả SA/MA/Matrix theo thứ tự position, stats mặc định `["base", "percent"]`
→ Nếu chỉ định cụ thể: chỉ thêm các câu đó theo đúng thứ tự user nhập

### Step 5 — Run pipeline (table-only)

Since `data/` already exists from Step 3, always run table-only:

```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    v1
```

**Force re-ingestion** (after fetching new data from QMe):
```bash
python run_pipeline.py \
  --mcp-dir         output/SURVEY_NAME/mcp \
  --export-csv      output/SURVEY_NAME/mcp/data_export.csv \
  --output-dir      output/SURVEY_NAME \
  --version         vX \
  --force-ingestion
```

> ⚠️ **NEVER recreate or rewrite `run_pipeline.py`** — it is always available.
> - Local dev (editable install): use `python run_pipeline.py`
> - PyPI install: use `surveyflow-run` (same arguments)

Nếu không dùng được `run_pipeline.py` (môi trường web/sandbox), chạy inline:
```python
import json, os, pathlib
from surveyflow import Pipeline, PipelineConfig

BASE       = pathlib.Path(os.getcwd())
output_dir = BASE / "output" / "SURVEY_NAME"
data_dir   = output_dir / "data"

# Ingestion-only (no datatable_config → table step is skipped):
from surveyflow.steps.ingestion.export_parser import parse_export_csv
mcp_dir    = output_dir / "mcp"
definition = json.load(open(mcp_dir / "definition.json", encoding="utf-8"))
export_df  = parse_export_csv(mcp_dir / "data_export.csv")
result = Pipeline(PipelineConfig(
    definition   = definition,
    export_df    = export_df,
    output_dir   = str(output_dir),
    data_dir     = str(data_dir),
)).run()
print("rawdata  :", result["rawdata_path"])
print("metadata :", result["metadata_path"])

# Table-only (after datatable.json is created):
result = Pipeline(PipelineConfig(
    output_dir       = str(output_dir),
    data_dir         = str(data_dir),
    skip_ingestion   = True,
    version          = "v1",
    datatable_config = str(output_dir / "datatable" / "datatable.json"),
)).run()
print("datatable:", result["datatable_path"])
print("rows     :", result["rawdata"].shape[0])
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
1. Read current `output/SURVEY_NAME/datatable/datatable.json`
2. Modify it according to user request
3. Save `datatable.json`
4. Run pipeline with new version (increment v1→v2→v3…) — table-only, no `--mcp-dir` needed

---

## datatable.json structure

`datatable.json` is an **array** — mỗi item là 1 table config độc lập, sinh ra các sheets riêng trong cùng 1 file xlsx.

**Kiểu item trong array:**
- `{ "type": "datatable", ... }` → bảng chéo thông thường (hoặc bỏ qua `type`, default là datatable)
- `{ "_custom_defs": [...] }` → khối định nghĩa filter dùng chung (không sinh sheet, chỉ dùng để tham chiếu)

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

Sheet tab name = `{sub_title} - {sheet}` → ví dụ `"General - Count"`, `"General - Pct"`, `"General - Sig"`.

### Banner rules
- Always include `{ "label": "Total", "filter": null }` as first entry
- `value` = single integer code from rawdata.csv
- `values` (plural) = list of codes to group together (e.g. age ranges)
- **Question reference: use the question label** (e.g. `"Q10"`) — this is the `label` field in metadata.json
- To find choice codes → read `output/SURVEY_NAME/data/metadata.json` → find question → inspect `choices_i18n`
- **MA questions are supported as banner** — each code becomes a column; respondents can appear in multiple columns

#### show_total — thêm cột Total trước các sub-groups

Thêm `"show_total": true` vào banner entry để tự động thêm cột "Total" (tất cả respondents có answer câu đó) ngay trước các sub-group columns:

```json
{
  "label": "Area",
  "question": "S3a",
  "groups": [
    { "label": "Hà Nội",   "value": 1 },
    { "label": "HCM",      "value": 2 }
  ],
  "show_total": true
}
```

Nếu **bất kỳ** banner entry nào có `show_total: true`, cột Total chung ở đầu sẽ bị ẩn (mỗi group đã có Total riêng).

`show_total` cũng dùng được trong mảng `levels` để thêm Total ở từng cấp nested.

### _custom_defs + custom_ref — reusable filter groups

Khi nhiều bảng (table items) cùng dùng 1 bộ filter giống nhau (ví dụ: User groups theo brand sử dụng), định nghĩa 1 lần trong `_custom_defs` rồi tham chiếu qua `custom_ref`.

**Bước 1 — Định nghĩa** (item đầu tiên trong array, không có `type`):
```json
{
  "_custom_defs": [
    {
      "label": "Users",
      "choices": [
        {
          "code": 1,
          "label": "Brand A users",
          "filter": { "question": "S11b", "codes": [1, 2] }
        },
        {
          "code": 2,
          "label": "Brand B users",
          "filter": { "and": [
            { "question": "S11b", "codes": [3, 4] },
            { "question": "S4",   "codes": [2, 3] }
          ]}
        }
      ]
    }
  ]
}
```

**Bước 2 — Tham chiếu** trong `banner` của bất kỳ table item nào:
```json
{ "type": "custom_ref", "ref": "Users", "label": "Users" }
```

Cũng dùng được bên trong mảng `levels` để tạo nested header:
```json
{
  "label": "Gender",
  "question": "S5",
  "groups": [{ "label": "Nam", "value": 1 }, { "label": "Nữ", "value": 2 }],
  "levels": [
    { "type": "custom_ref", "ref": "Users", "label": "Users" }
  ]
}
```
→ Header 3 cấp: `Gender / Nam|Nữ / Brand A|Brand B`

**Filter syntax trong `_custom_defs`:**
- Leaf: `{ "question": "Qx", "codes": [1,2,3], "op": "any" }` — `op` default là `"any"` (chọn ít nhất 1 code)
- AND: `{ "and": [ ...leaf... ] }`
- OR: `{ "or": [ ...leaf... ] }`

### banner_matrix — Matrix rows as nested banner columns

Use `banner_matrix` to add a matrix question's rows (e.g. brands) as the innermost level of every banner column. When active, stub matrix questions switch to **paired mode** — instead of expanding by sub-question, they show choice distributions matched to each brand column.

```json
"banner_matrix": {
  "label": "Brand",
  "question": "Q17"
}
```

Without `groups`: all rows from `choices_i18n.rows` become individual columns.

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

**Groups rules:**
- `row_code` (string): single brand row → paired mode reads `{q_col}_r{code}` per column
- `row_codes` (list): grouped brands → stacked mode sums counts across all `{q_col}_r{rc}` columns
- Mix of both is allowed in the same `groups` array

**Header levels produced:**
- Total column + each brand: `Total / Brand` (2 levels)
- Store Type column + each brand: `Store Type / Sub-group / Brand` (3 levels)

### Stub rules
- One entry per question
- `stats` options: `"base"`, `"percent"`, `"t2b"`, `"b2b"`, `"mean"`, `"std"`, `"se"`
- Supported answer types: `SA`, `MA`, `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`
  - Matrix questions automatically expand into one block per row (sub-question)
  - When `banner_matrix` is active, matrix questions use paired mode instead
- Excluded automatically: `FT`, `NUM`, `instruction`, `user-name`, `user-phone`, `reward`, `record`
- **Stub order follows datatable.json** — pipeline outputs questions in the order they appear

#### row_group — nhóm matrix questions theo shared row headers

Dùng `row_group: true` khi muốn nhóm nhiều câu matrix có chung `choices_i18n.rows` (exact match). Pipeline render 1 section header cho mỗi row label, với sub-blocks cho từng câu bên dưới.

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
1. Tất cả items phải là matrix questions (`Matrix_SA`, `Matrix_MA`, `Matrix_NUM`)
2. Tất cả items phải có cùng `choices_i18n.rows` (exact match)
3. Không được mix non-matrix questions vào trong `row_group`

#### Sub-question reference — Q{label}_r{n}

Để lấy 1 row cụ thể của matrix làm flat stub (ví dụ: chỉ 1 brand):

```json
{ "question": "Q17_r10",      "label": "SK ZIC — satisfaction",    "stats": ["base", "percent"] },
{ "question": "Q14_Freq_r10", "label": "SK ZIC — visit frequency", "stats": ["base", "percent"] }
```

- `r10` = sub-question có `row_index == 10` trong metadata
- Dùng được với mọi matrix type, ngoài `row_group`

### Tables rules
- Khi khởi tạo lần đầu: tạo đúng theo lựa chọn của user ở Step 4
- Sau khi đã có datatable.json: không thay đổi `tables` trừ khi user yêu cầu
- `enabled: false` → sheet đó bị bỏ qua khi chạy
- **`decimal`** — số chữ số thập phân cho cột percentage: `0` → "0%" (mặc định), `1` → "0.0%", `2` → "0.00%"
- **Sig test config** đặt **trên từng sheet** (không dùng block `significance_test` cấp trên nữa):
  - `show_sig: true` → bật sig test cho sheet đó
  - `levels` (optional, default `[90, 95]`) → danh sách confidence levels
  - `method` (optional, default `"independent"`) → `"independent"` (Welch's) hoặc `"related"` (paired)
- Mapping lựa chọn → sheets:
  - Count only → `[{ "sheet": "Count", "cell_content": "count",      "show_sig": false, "enabled": true }]`
  - Pct only   → `[{ "sheet": "Pct",   "cell_content": "percentage", "show_sig": false, "enabled": true, "decimal": 0 }]`
  - Sig        → `[{ "sheet": "Sig",   "cell_content": "percentage", "show_sig": true,  "levels": [90, 95], "method": "independent", "enabled": true }]`
  - Tất cả    → cả 3 sheets: Count + Pct + Sig

> ⚠️ Block `significance_test` ở cấp trên (nếu còn trong file cũ) bị **bỏ qua hoàn toàn**. Sig config chỉ đọc từ `tables[]`.

---

## Auto-generate datatable.json rules
> ⚠️ **Không dùng auto-generate khi khởi tạo lần đầu.** Luôn hỏi user theo Step 4.
> Auto-generate chỉ dùng khi user nói rõ: *"tự động tạo"* hoặc *"auto"*.

**Banner:** Pick singlechoice (SA) questions that look demographic:
- Gender, Age, Location, Income, Marital status, Occupation
- Max 4-5 banner groups + Total

**Stub:** Include all codeable questions sorted by position:
- `SA`, `MA`, `Matrix_SA`, `Matrix_MA`
- Default stats: `["base", "percent"]`
- Skip: `FT`, `NUM`, `instruction`, `user-name`, `user-phone`, `reward`, `record`

---

## Output structure
```
output/SURVEY_NAME/
├── mcp/                      ← MCP raw files (fetched from QMe)
│   ├── definition.json
│   └── data_export.csv       ← assembled from read_survey_data_file chunks
├── data/                     ← rawdata.csv + metadata.json (generated from mcp/, reused)
│   ├── rawdata.csv
│   └── metadata.json
├── datatable/                ← Claude manages this file
│   └── datatable.json
├── v1/                       ← datatable.xlsx only
│   └── datatable.xlsx
├── v2/                       ← after user requests changes
│   └── datatable.xlsx
└── v3/
    └── datatable.xlsx
```

---

## Common user requests → datatable.json changes

| User says | Claude does |
|---|---|
| "thêm income vào banner" | Add income question to `banner` array |
| "bỏ Q15 khỏi stub" | Remove Q15 entry from `stub` array |
| "thêm mean std cho Q36" | Add `"mean"`, `"std"` to Q36's `stats` |
| "thêm tất cả câu vào stub" | Add all codeable questions to `stub` |
| "tắt sig test" | Xoá `show_sig: true` hoặc set `"show_sig": false` trên sheet Sig |
| "bật sig test 90% và 95%" | Thêm `"show_sig": true, "levels": [90, 95]` vào sheet Sig |
| "hiện 1 chữ số thập phân" | Thêm `"decimal": 1` vào sheet Pct |
| "thêm total cho từng banner group" | Thêm `"show_total": true` vào banner entry tương ứng |
| "chỉ chạy 1 sheet percentage" | Modify `tables` to keep only Pct sheet |
| "nhóm Q13/Q14/Q17 theo brand" | Dùng `row_group: true` với các câu đó trong stub |
| "thêm SK ZIC riêng cho Q17" | Thêm `Q17_r{n}` sub-question ref vào stub |
| "tạo filter group dùng chung" | Tạo item `_custom_defs` đầu array, định nghĩa `choices` với `filter` |
| "thêm user groups vào banner" | Thêm `{ "type": "custom_ref", "ref": "DefName" }` vào banner |
| "user groups × area nested" | Dùng `levels: [{ "type": "custom_ref", ... }]` trong banner entry |
| "brand làm header, tất cả brands" | Thêm `banner_matrix: { question: "QX" }` (không có groups) |
| "brand header, Castrol và Shell riêng, nhóm International" | Thêm `banner_matrix` với `groups` mix `row_code`/`row_codes` |
| "bảng bình thường + bảng matrix brand" | Tạo 2 items trong array: 1 không có banner_matrix, 1 có |
| "refresh data / lấy data mới" | Re-fetch (prepare→poll→read) → overwrite `mcp/data_export.csv` → re-run với `--export-csv` + `--force-ingestion` |

---

## Profile status
Default: `approved` only.
- Pipeline CLI: `--profile-status approved,pending`
- MCP tool call: `profile_status=["approved"]` (array, not string)

---

## QMe MCP tool reference

### get_survey_definition
```
Required: survey_id (integer)
```

### prepare_survey_data_file
```
Required: survey_id     (integer)
Optional: format        (string "code" | "text", dùng "code")
          force_refresh (boolean, default false — dùng cache TTL)
```
Returns: `{ job_id, status, expires_at, files[], ... }`

### get_survey_data_file_status
```
Required: job_id (string)
```
Returns: `{ status ("pending"|"processing"|"ready"|"error"), retry_after_seconds, ... }`  
Poll mỗi `retry_after_seconds` cho đến khi `status == "ready"`.

### read_survey_data_file
```
Required: job_id  (string)
Optional: file    (string "data" | "definition", default "data")
          offset  (integer, default 0)
          limit   (integer, default 500)
```
Returns: `{ rows/csv, pagination: { has_more, next_offset, rows_remaining } }`  
Pagination: lặp lại với `offset += limit` cho đến khi `pagination.has_more == false`.

### search_surveys (dùng khi tìm survey_id)
```
Optional: query (string), status (array), offset, limit
```
