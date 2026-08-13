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

**Fetch / Upload zip → Ingestion → Quality check (optional) → Datatable → Appendix slides**

> ⚠️ **Behavior rule — BẮT BUỘC:**
> Claude **KHÔNG tự suy đoán và thực hiện các bước ngoài skill**.
> Nếu không rõ bước nào, hoặc gặp issue → **hỏi lại user trước** khi tiếp tục.
> Không tự ý skip bước, đoán tên survey, hoặc chạy pipeline khi chưa đủ thông tin.

---

## How to start

```
/fieldcheck-dp run pipeline VN8966
chạy survey VN8966
làm bảng cho VN8894 - Express
fetch data VN8894
chạy quality check VN8966
```

Nếu user không nói tên survey → hỏi ngay: *"Bạn muốn chạy survey nào?"*

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
├── mcp/                    ← raw MCP files
│   ├── definition.json
│   └── data_export.csv
├── data/                   ← rawdata.csv + metadata.json
├── datatable/
│   └── datatable.json
├── quality/                ← quality check output (optional)
├── v1/
│   ├── datatable.xlsx
│   ├── chart_data.json
│   └── slides.pptx
└── v2/, v3/, ...
```

---

## Workflow A — First run

### Step 0 — Show progress tracker

```
📋 Pipeline: SURVEY_NAME
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⏳ 1. Tìm survey         — đang tìm...
⬜ 2. Lấy data
⬜ 3. Ingestion
⬜ 4. Chọn banner / stub
⬜ 5. Chạy bảng
⬜ 6. PPTX appendix      (tuỳ chọn)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

Icons: `⏳` đang chạy · `✅` xong · `⬜` chờ · `❌` lỗi. Always end each update with a next-action hint.

---

### Step 1 — Find survey

```
search_surveys(query="SURVEY_NAME")
```

Note the `survey_id`. If multiple results → show list, ask user to confirm.

---

### Step 2 — Lấy data

**Nếu `output/SURVEY_NAME/mcp/` đã có `definition.json` + `data_export.csv`:**
→ Hỏi: *"Data đã có sẵn. Dùng data cũ hay lấy lại?"*

**Nếu chưa có data** → Hỏi user:
> *"Bạn muốn **fetch từ QMe**, **upload file zip** từ Fieldcheck, hay **upload file Excel**
> "Question"+"Data" 2-sheet?"*

- **Fetch từ QMe** → Step 2A
- **Upload zip** → Step 2B
- **Upload xlsx "Question"+"Data"** → Step 2C

> ⚠️ Fetch rules: `format="code"` always · Never use `get_survey_rows` · Write tool for JSON/CSV · `data_export.csv` encoding=`utf-8-sig`

**Step 2A — Fetch từ QMe:**

```
get_survey_definition(survey_id)  → output/SURVEY_NAME/mcp/definition.json

prepare_survey_data_file(survey_id, format="code", force_refresh=False)
  → job_id

get_survey_data_file_status(job_id)   ← poll every retry_after_seconds until status=="ready"
  → If stuck after 3+ polls: stop → tell user to export zip manually → switch to Step 2B

read_survey_data_file(job_id, offset=0,   limit=500)
read_survey_data_file(job_id, offset=500, limit=500)
... until pagination.has_more == false
→ Assemble all chunks → write to output/SURVEY_NAME/mcp/data_export.csv (utf-8-sig)
```

Tell user: `"✅ Fetch xong — {N} responses"`

**Step 2B — Upload zip:**

```python
import zipfile, os
with zipfile.ZipFile('uploaded.zip') as z:
    data_file = next((n for n in z.namelist() if n.startswith('code_retail_report_')), None)
    if not data_file:
        raise ValueError("Không tìm thấy file code_retail_report_* trong zip")
    content = z.open(data_file).read().decode('utf-8-sig')
os.makedirs('output/SURVEY_NAME/mcp', exist_ok=True)
with open('output/SURVEY_NAME/mcp/data_export.csv', 'w', encoding='utf-8-sig') as f:
    f.write(content)
```

> Nếu `definition.json` chưa có: fetch bằng `get_survey_definition(survey_id)` trước.

**Step 2C — Upload xlsx "Question"+"Data":**

Khi user đính kèm 1 file `.xlsx` gồm đúng 2 sheet **"Question"** (danh sách câu hỏi dạng
phẳng) và **"Data"** (khớp định dạng `data_export.csv` chuẩn — header có ô đầu `"Approve"`)
→ dùng thẳng `--xlsx-input`, **không tự viết converter riêng**:

```bash
python run_pipeline.py \
  --xlsx-input "SURVEY_NAME.xlsx" \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --output-dir output/SURVEY_NAME \
  --version    v1
```

surveyflow (`surveyflow/steps/ingestion/flat_xlsx_import.py`) tự convert sang
`definition.json` + `data_export.csv`, tự suy luận SA/MA/FT/RANKING/Matrix_SA/multiplenumber,
tự loại field interviewer-only và field ảnh (label kết thúc `_PHOTO`). **Luôn đọc phần
`WARNING` in ra console** sau khi chạy — liệt kê mọi cột bị skip/giả định, cần kiểm tra lại
nếu survey có cấu trúc khác biệt lớn (converter mới verify với SA/MA/FT/RANKING/Matrix_SA/
multiplenumber — chưa gặp Matrix_MA, NUM, hay synthetic type gender/area/personal-income).

> Sau ingestion vẫn cần chạy Step 3a/3b/3c như bình thường (title_i18n, scale_class,
> mean/factor/T2B/NPS) — converter chỉ tạo đúng structure, không tự phân loại/tóm tắt.

---

### Step 3 — Run ingestion

```bash
python run_pipeline.py \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --export-csv output/SURVEY_NAME/mcp/data_export.csv \
  --output-dir output/SURVEY_NAME
```

Output: `data/rawdata.csv` + `data/metadata.json`

After ingestion, tell user:

```
✅ Ingestion xong — {N} rows, metadata.json sẵn sàng.

💡 Bạn có muốn chạy quality check trước khi tạo bảng không?
   Gõ "chạy quality" hoặc "bỏ qua" để tiếp tục.
```

**Chỉ chạy quality check khi user xác nhận.**

Sau đó (không cần đợi quality check) → luôn chạy **Step 3a — Tóm tắt title_i18n** và
**Step 3c — Phân loại câu SA/Matrix_SA** trước khi sang Step 4 (thứ tự giữa 3a/3c không quan
trọng, độc lập nhau).

---

### Step 3a — Tóm tắt tiêu đề câu hỏi (title_i18n)

`parse_metadata()` (code) đã tự ghi field **`title_i18n: null`** vào mọi câu hỏi trong
`metadata.json` (và mọi entry trong `sub_questions` của câu matrix) — code không tự tóm tắt được,
chỉ Claude mới điền nội dung thật. **Bỏ qua bước này nếu mọi câu đã có `title_i18n` khác `null`**
(idempotent, chỉ chạy 1 lần sau ingestion, giống Step 3c với `scale_class`).

1. Đọc `question_i18n` (`vi` + `en`) từng câu.
2. Viết tiêu đề ngắn gọn (5-10 từ), giữ đúng ý, bỏ phần kỹ thuật/lặp ("(Multiple answers
   allowed)", placeholder `{QD9/802773/Selected}`...) — làm cho cả `vi` lẫn `en`:
   ```json
   "title_i18n": { "vi": "Lý do chọn thương hiệu", "en": "Reasons for brand choice" }
   ```
3. **Matrix sub_questions** — không có sẵn `question_i18n` riêng, ghép tiêu đề câu cha +
   `row_label` của dòng đó: `{ "en": "Brand satisfaction — Brand A" }`.
4. Ghi đè `metadata.json`.

Chạy `run_pipeline.py`/`surveyflow-run` thẳng từ terminal (không qua Claude) → `title_i18n` giữ
`null` mãi mãi, đúng thiết kế (không phải lỗi). `title_i18n` giờ **đã nối vào chart_data.json**:
khi table step tạo `chart_data.json`, mỗi câu có field `"title"` (ngay sau `"question"`) —
lấy từ `"title"` của stub trong `datatable.json` nếu có, không thì tự fallback sang
`title_i18n[lang]` của câu đó trong `metadata.json` (giống hệt cách `"label"` fallback sang
`question_i18n`). Nhờ vậy chỉ cần làm Step 3a **một lần** sau ingestion là mọi slide appendix
sau này tự có tiêu đề tóm tắt, không cần set `"title"` thủ công cho từng stub entry.

---

### Step 3b — Quality check (user confirms)

Chạy khi user đồng ý, hoặc khi user nói: *"chạy quality", "kiểm tra data", "check lỗi routing"*.

```bash
python run_pipeline.py --output-dir output/SURVEY_NAME --run-quality
```

Read `quality/quality_report.json` và present:

```
📊 Quality Report — SURVEY_NAME
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Tổng respondents  : {total_respondents}
Bị flag           : {flagged_count} ({pct:.1f}%)

Loại vi phạm:
  ❌ missing          {n}
  ⚠️  routed_missing   {n}
  🔍 extraneous       {n}
  💥 contradiction    {n}

Câu bị flag nhiều nhất (top 5):
  {Q_label} — {N} lần
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📄 Chi tiết: output/SURVEY_NAME/quality/flagged_profiles.csv
```

If `flagged_count == 0` → `"✅ Không có vi phạm nào."`

Sau đó hỏi: *"Bạn muốn xem chi tiết câu nào, hay tiếp tục chạy bảng?"*

**Drill-down:** "Xem chi tiết Q5" → filter violations by question · "Profile bị lỗi nhiều nhất" → group by profile_id top 10 · "Chỉ xem contradiction" → filter by type.

---

### Step 3c — Phân loại câu SA VÀ Matrix_SA (Ordinal / Nominal)

> ⚠️ **Quét CẢ `SA` lẫn `Matrix_SA` — lỗi thực tế đã xảy ra**: lọc cứng `answer_type == "SA"`
> sẽ bỏ sót toàn bộ Matrix_SA (VD Q17/Q15/Q11/Q6_A-kiểu-câu), khiến chúng không bao giờ có
> `scale_class` → không bao giờ tự thêm mean/T2B/NPS ở Step 3d. Matrix_SA phải ghi
> `scale_class` vào **CẢ entry cha lẫn từng entry trong `sub_questions`** (chỉ ghi entry cha
> không đủ — appendix PPTX khớp theo mã sub-question như `A4_r1`, không phải mã câu cha).

Chạy tự động ngay sau ingestion — không cần hỏi user (giống Step 3, không phải hành động
phá huỷ dữ liệu). **Bỏ qua bước này nếu `metadata.json` đã có field `scale_class`** (đã
phân loại từ lần chạy trước).

Claude đọc từng câu `answer_type: "SA"` **và** `answer_type: "Matrix_SA"` trong `metadata.json`
(question_i18n + choices_i18n, với Matrix_SA thì đọc `choices_i18n.columns`) và tự phân loại —
dùng khả năng đọc hiểu ngữ cảnh, **không phải rule cố định trong code**:

| Loại câu hỏi | Tên nên dùng | Lưu vào `scale_class` |
|---|---|---|
| SA/Matrix_SA có thang đo (rating, mức độ đồng ý, purchase intent, tần suất, willingness-to-pay, time-since...) | SA Scale / SA Likert Scale | `"Ordinal"` |
| SA/Matrix_SA phân loại thường (demographic, brand, tradeoff, typology, awareness...) | SA Categorical / SA Nominal | `"Nominal"` |

> ⚠️ **Kiểm tra chiều thang đo — ghi thêm `scale_high_code` cho thang ĐẢO NGƯỢC.** Đừng giả định
> code lớn = điểm cao. Một số thang (đặc biệt **tần suất**) đánh số ngược: code 1 = mức cao nhất
> ("Hầu như mỗi ngày"), code lớn nhất = thấp nhất ("Ít hơn 1 lần/tháng"). Khi phân loại 1 câu là
> `Ordinal` mà **code nhỏ nhất là đầu "cao nhất"** của thang → ghi thêm field **`scale_high_code`** =
> code đầu cao đó (VD `1`). Step 3d dùng field này để tính mirror factor + nhóm NPS đúng chiều; thiếu
> nó → mean/T2B/NPS sai chiều. **Bỏ qua** field này nếu code lớn nhất đã là đầu "cao nhất" (trường
> hợp thường gặp — satisfaction/agreement/purchase intent). Matrix_SA: ghi vào **cả entry cha lẫn
> từng `sub_questions`**, giống `scale_class`.

```python
import json
path = 'output/SURVEY_NAME/data/metadata.json'
with open(path, encoding='utf-8') as f:
    meta = json.load(f)
qs = meta['questions']

# Claude đọc question_i18n + choices_i18n (Matrix_SA: choices_i18n.columns) của từng câu
# SA/Matrix_SA chưa có scale_class, tự quyết định "Ordinal" hoặc "Nominal" rồi gán:
# qs[qid]['scale_class'] = "Ordinal"   # hoặc "Nominal"
# Thang ĐẢO NGƯỢC (code nhỏ nhất = đầu cao nhất, VD thang tần suất) → ghi thêm scale_high_code:
# qs[qid]['scale_high_code'] = 1       # code đầu "cao nhất"; BỎ QUA nếu code lớn nhất đã là đầu cao
# Matrix_SA: PHẢI gán CẢ scale_class LẪN scale_high_code cho từng sub_questions entry, không chỉ cha:
# for sub in qs[qid].get('sub_questions', {}).values():
#     sub['scale_class'] = qs[qid]['scale_class']
#     if 'scale_high_code' in qs[qid]: sub['scale_high_code'] = qs[qid]['scale_high_code']

with open(path, 'w', encoding='utf-8') as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)
```

Sau khi ghi xong, **luôn báo cho user** số câu mỗi loại + liệt kê tên từng câu (dùng
`label`, không phải `question_id`):

```
📐 Phân loại câu SA/Matrix_SA — SURVEY_NAME
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Ordinal ({n} câu): {label1}, {label2}, ...
Nominal ({m} câu): {label1}, {label2}, ...
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
👉 Câu Ordinal sẽ tự động thêm "mean" vào stats khi tạo datatable.json.
Bạn có muốn đổi lại loại của câu nào không?
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

> Ví dụ Ordinal: thang hài lòng (1–5), tần suất (Never→Always), mức độ đồng ý, purchase
> intent (Definitely won't buy→Definitely will buy), willingness-to-pay bands, time-since bands.
> Ví dụ Nominal: giới tính, khu vực — kể cả **Age/Income dạng bracket** (banner/demographic
> var, không tính mean dù có thứ tự), thương hiệu, loại sản phẩm, tradeoff/typology choices.

Nếu user đổi loại một câu ("đổi Q10 thành Nominal") → sửa `scale_class` trong
`metadata.json` cho câu đó, ghi đè lại file, xác nhận với user.

---

### Step 3d — Detect Ordinal (tự động thêm stats/factor/group codes)

Khi thêm câu `scale_class: "Ordinal"` vào `stub` (Step 4), tự động bổ sung thêm — không cần
user yêu cầu riêng (bỏ qua meta code 98/99 trong mọi phép tính):

1. **Mọi câu Ordinal → thêm `"mean"` vào `stats` + LUÔN thêm `"factor"` cho từng choice**
   (bắt buộc — thiếu factor dù chỉ 1 code sẽ khiến pipeline âm thầm trả mean=0.0, không có
   fallback tính trên code gốc). Không đảo ngược → `factor(code) = code` (identity). Đảo
   ngược (`scale_high_code` = code nhỏ nhất) → `factor(code) = code_max + code_min - code`.
   **Chỉ 1 format được hỗ trợ:** per-choice `"choices": [{"code":.., "factor":..}]` — KHÔNG
   có top-level `"factors"`/`"mean_factor"` (đã bị bỏ hẳn khỏi pipeline) hay `"factor"` số ít
   (chưa từng được hỗ trợ) — cả 2 đều dẫn tới mean âm thầm = 0.0, đã xảy ra thực tế 2 lần.
   Matrix_SA: code lấy từ `choices_i18n.columns`, **KHÔNG BAO GIỜ** từ `.rows` (rows = item
   được đánh giá như brand, không phải thang đo) — cũng là bug thực tế đã xảy ra.

   **⚠️ NGOẠI LỆ — câu tần suất (frequency): KHÔNG áp dụng mục 1 tự động.** Thang tần suất
   (VD "Hàng ngày/Vài lần tuần/Hiếm khi/Không bao giờ") thường không có 1 factor "đúng" hiển
   nhiên (khoảng cách giữa các mức không đều, hoặc có mức bất thường phá vỡ thứ tự). Gặp câu
   Ordinal dạng tần suất → **hỏi user trước**: *"Câu {label} là thang tần suất. Bạn muốn tự
   nhập factor cho từng mức hay để pipeline coi câu này là Nominal (không tính mean)?"*
   - User cung cấp factor → dùng đúng giá trị đó, giữ `scale_class: "Ordinal"`, vẫn áp
     T2B/B2B/NPS bình thường theo range (mục 2-3 dưới đây).
   - User chọn bỏ qua → sửa `scale_class` câu đó thành `"Nominal"` trong `metadata.json`
     (Matrix_SA: cả entry cha + từng `sub_questions`), stub chỉ giữ `["base","percent"]`.
2. **Thang 1–5** → thêm 2 group entry vào cuối mảng `choices` (KHÔNG thêm `"t2b"`/`"b2b"` vào
   `stats` — 2 stat name này không còn tồn tại, dùng format cũ sẽ bị bỏ qua lặng lẽ):
   ```json
   { "label": "T2B", "codes": [4, 5], "type": "combine" },
   { "label": "B2B", "codes": [1, 2], "type": "combine" }
   ```
   2 code gần đầu "cao nhất" theo `scale_high_code` → T2B; 2 code gần đầu "thấp nhất" → B2B.
   `stats` chỉ cần có `"percent"` — group tự render kèm theo.
3. **Thang 1–10 hoặc 0–10** → thêm `"nps"` vào `stats` + `choices` groups:
   ```json
   "choices": [
     { "type": "promoters",  "codes": [9, 10],              "label": "Promoters" },
     { "type": "passive",    "codes": [7, 8],                "label": "Passives" },
     { "type": "detractors", "codes": [0, 1, 2, 3, 4, 5, 6], "label": "Detractors" }
   ]
   ```
   (thang 1–10 không có code 0 → Detractors bỏ code 0.) Đảo ngược → nhóm theo khoảng cách
   tới `scale_high_code`, không theo trị số tuyệt đối.

Thang đo khác 1–5 và 1–10/0–10 → chỉ thêm `"mean"` (mục 1).

**`"type"` hợp lệ trong group entry** (viết thường): `"combine"` (mặc định, gộp 1 hàng %) ·
`"netted"` (gộp 1 hàng % tổng + các hàng % riêng từng code bên dưới) · `"promoters"`/`"detractors"`
(chỉ 2 giá trị này được stat `"nps"` đọc để tính điểm). `"passive"` chỉ có ý nghĩa trình bày,
tương đương `"combine"`.

---

### Step 4 — Design datatable.json

**Nếu user đã chỉ định yêu cầu** → tạo `datatable/datatable.json` trực tiếp, sang Step 4b.

**Nếu chưa chỉ định** → hỏi tuần tự 3 câu (chờ user trả lời xong mỗi câu mới hỏi câu tiếp):

**Câu 1 — Banner:**
Đọc `metadata.json`, liệt kê các câu SA/MA (thường là câu demographics):
```
Banner gồm những câu nào? (Total luôn có sẵn)
Ví dụ: S3 (Gender), S5 (Age), S7 (Income)
Nhập số câu cách nhau bằng dấu phẩy:
```
→ Lấy choice codes từ `metadata.json` để tạo `groups`.

**Câu 2 — Stub:**
Liệt kê tất cả câu SA/MA/Matrix từ `metadata.json`:
```
Stub gồm những câu nào?
- Nhập "all" để lấy tất cả
- Hoặc nhập số câu: Q1, Q5, Q8...
```
→ Câu SA **hoặc Matrix_SA** có `scale_class: "Ordinal"` trong `metadata.json` → xem **Step 3d**
  để tự động thêm `"mean"` + **factor** (bắt buộc, mọi choice) + group T2B/B2B trong `choices`
  cho thang 1-5, `"nps"` cho thang 1-10/0-10 — không cần user yêu cầu riêng. Câu Nominal hoặc
  chưa có `scale_class` thì giữ mặc định `["base", "percent"]`.
→ Câu `NUM` → LUÔN tự động thêm `"mean"` + `"num_quantile": 4`, không cần hỏi.
→ Câu `multiplenumber` → LUÔN tự động thêm `"mean"`, không cần hỏi.

**Câu 3 — Title:**
```
Tiêu đề bảng? (Enter để dùng mặc định: "SURVEY_NAME - Data Table")
```

Tạo `output/SURVEY_NAME/datatable/datatable.json` với **default tables: Count + Pct** (không có Sig — xem Tables rules).

**Helper script** — `group_numeric.py` (chỉ dùng cho câu `NUM`, khi cần xem trước bin trước khi
chốt cách group). Banner/stub choice codes thì đọc thẳng `metadata.json`, không có script riêng:
```bash
python scripts/group_numeric.py output/SURVEY_NAME/data/rawdata.csv output/SURVEY_NAME/data/metadata.json --question S3_1
```
> `scripts/` = tương đối với thư mục skill. Khi làm việc trong repo surveyflow, path đầy đủ là
> `.skill_work/fieldcheck-dp/scripts/group_numeric.py`.
Đề xuất range/bucket từ phân bố giá trị thật (bin width "đẹp" — 1/2/5/10/20/25/50/100 × 10^n),
in sẵn `"choices"` (group + hidden) để paste vào stub. Flag: `--width N` ép width cụ thể ·
`--bins N` đổi số bin mục tiêu (mặc định 6) · `--quantile N` chuyển sang bin ~bằng nhau về số
respondent, khớp đúng những gì `"num_quantile": N` sinh ra lúc chạy table (dùng để chọn N).

---

### Step 4b — Language check

Trước khi chạy pipeline, kiểm tra metadata.json xem survey có mấy ngôn ngữ:

```python
import json
with open('output/SURVEY_NAME/data/metadata.json') as f:
    meta = json.load(f)
# metadata.json KHÔNG có key "languages" — suy ra từ keys của question_i18n
langs = sorted({k for q in meta['questions'].values() for k in q.get('question_i18n', {})})
print('Languages:', langs)
```

- **1 ngôn ngữ** → dùng luôn, không hỏi user (ví dụ: `--lang vi`)
- **Nhiều ngôn ngữ** → hỏi:
  > *"Survey có {langs}. Bạn muốn output ngôn ngữ nào?"*
  → User chọn → dùng `--lang {choice}` khi chạy Step 5

---

### Step 5 — Run pipeline

```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    v1 \
  --lang       vi
```

Force re-ingestion after new fetch:
```bash
python run_pipeline.py \
  --mcp-dir         output/SURVEY_NAME/mcp \
  --export-csv      output/SURVEY_NAME/mcp/data_export.csv \
  --output-dir      output/SURVEY_NAME \
  --version         vX \
  --lang            vi \
  --force-ingestion
```

> ⚠️ NEVER recreate or rewrite `run_pipeline.py`.

After pipeline:
```
✅ Datatable xong!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 File   : output/SURVEY_NAME/v1/datatable.xlsx
📊 Sheets : General - Count  |  General - Pct
👥 Rows   : {N} respondents
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bạn muốn thay đổi gì không? (thêm banner, thêm câu, bật sig test, tạo PPTX...)
```

---

### Step 6 — PPTX Chart Appendix (tuỳ chọn)

Khi user nói *"tạo appendix / chạy slides / tạo PPTX"* → confirm rồi chạy.

**Chọn table** — appendix render **mọi table** có `"is_appendix": true` (KHÔNG còn theo prefix
`sub_title` nữa — `sub_title` giờ là tên tự do, dùng luôn làm tên group, VD "Hảo Hảo", "Total"),
gộp thành **1 file .pptx**.

- **Chỉ 1 table `is_appendix: true`** → chạy thẳng, không hỏi, không chia group.
- **≥2 table** → mỗi group: 1 PowerPoint Section (Slide Sorter/Outline) bọc đúng slide của group
  đó, KHÔNG chèn slide riêng cho tên group. Style/logo áp dụng cho cả deck, lấy từ group đầu tiên.
- **Chưa table nào đánh dấu** → hỏi user **một lần** (không hỏi lại mỗi lần chạy, giống
  format/logo — lưu thẳng vào `datatable.json`):
  1. `surveyflow-pptx <chart_data.json> --list-groups` → in JSON **mọi table** kèm `is_appendix`
     hiện tại (không tạo pptx).
  2. Hỏi: *"Bạn muốn chạy appendix cho bảng nào? {sub_title1}, {sub_title2}... (chọn 1/nhiều
     tên, hoặc 'All')"*
  3. Theo trả lời → **sửa `datatable.json`**: set `"is_appendix": true` cho table được chọn
     ("All" = mọi table); table nào trước đó `true` nhưng lần này không chọn lại → set `false`.
  4. Chạy `surveyflow-pptx`/`--appendix` bình thường — không cần `--tables` nữa (field đã quyết
     định sẵn).

```bash
# Kèm pipeline:
python run_pipeline.py --output-dir output/SURVEY_NAME --version vX --lang vi --appendix

# Riêng (sau khi đã có chart_data.json):
surveyflow-pptx output/SURVEY_NAME/vX/chart_data.json output/SURVEY_NAME/vX/slides.pptx
```

Options: `--list-groups` (in JSON mọi table + `is_appendix` rồi thoát, không tạo pptx) ·
`--tables idx1,idx2` (override 1 lần: chạy đúng các table_index đã `is_appendix:true`, không sửa
file) · `--table N` (ghi đè bằng đúng 1 table_index, bỏ qua toàn bộ cơ chế group) ·
`--start-page N` (số trang bắt đầu) · `--format general|default` (ghi đè `appendix_format` 1
lần, không lưu) · `--logo acecook|none|path/to/logo.png` (ghi đè `appendix_logo` 1 lần, chỉ với
`--format default`) · `--default-section-label "08 | PHỤ LỤC"` (ghi đè tag góc trên-trái)

**Format** — mặc định luôn dùng `"default"` (style công ty), **không cần hỏi**. Chỉ khi user yêu cầu
rõ style thuần surveyflow (không branding) → ghi `"appendix_format": "general"` vào table item trong
`datatable.json`. `"default"`: slide layout + placeholder title/footer/số trang riêng của template
công ty, font Segoe UI, bảng màu riêng (5 màu thật của brand + 15 sắc tint/shade suy ra, tổng 20 màu
— tránh trùng màu khi câu có >5 lựa chọn), chỉ 2 layout (donut+stack cho Ordinal; 3-cột bar ngang
dùng chung cho MA và SA-Nominal nhiều lựa chọn — chưa có style cột dọc riêng). Tag góc trên-trái tự
đổi theo `--lang` của lần chạy table gần nhất: "PHỤ LỤC" (vi) / "APPENDIX" (en) — không cần cấu hình.
Asset: `surveyflow/steps/appendix/appendix_templates/default_template.pptx` (đã xoá slide mẫu, không
dùng lại để re-extract) + `chart_templates_default/{bar,donut,stacked}.xml`.

**Logo khách hàng** (chỉ format `"default"`) — hỏi **một lần** nếu table chưa có `appendix_logo`:
> *"Logo khách hàng trong appendix? 1. Acecook (có sẵn) · 2. Khách khác (cần file ảnh) · 3. Không dùng (chỉ Q&Me)"*
→ 1 = bỏ qua field (mặc định `"acecook"`) · 2 = hỏi đường dẫn ảnh → ghi vào `"appendix_logo"` ·
3 = `"appendix_logo": "none"`. Logo Q&Me (trái) **không bao giờ** đổi; chọn `"none"` cũng xoá dấu gạch
dọc "|" giữa 2 logo (shape `QMeLogo`/`CustomerLogo`/`LogoDivider` trong slide master). Cả
`appendix_format` lẫn `appendix_logo` tự chảy qua `chart_data.json` → hỏi 1 lần, các lần chạy sau tự
dùng lại, KHÔNG hỏi lại.

**Tiêu đề slide tuỳ chỉnh** — stub entry có thể thêm field `"title"` (mặc định `null`). Thứ tự
fallback khi tạo `chart_data.json`: `"title"` của stub (nếu set) → `title_i18n[lang]` của câu đó
trong `metadata.json` (Step 3a, tự động, không cần làm gì thêm) → label gốc rút gọn (nếu cả 2 đều
null). Chỉ cần **ghi đè thủ công** field `"title"` khi muốn 1 tiêu đề KHÁC với `title_i18n` đã có
sẵn (VD label dài/khó đọc mà bản tóm tắt Step 3a chưa ưng ý) — đề xuất bản tóm tắt ngắn (5-10 từ)
→ user xác nhận → ghi vào `"title"` (KHÔNG tự ghi mà không hỏi, khác các auto-rule khác — tóm tắt
ngôn ngữ tự nhiên là chủ quan). Footer/Q-label cuối slide vẫn luôn hiện label gốc đầy đủ, `title`
chỉ thay tiêu đề lớn. **KHÔNG tự thêm tiền tố `"{sub_title} - "` (tên brand/group) vào `title`** —
PowerPoint Section đã hiển thị tên group khi deck gộp nhiều table, lặp lại trong từng title là dư
thừa (bug thực tế đã xảy ra ở VN8971, đã dọn lại). Matrix/ranking: 1 `title` trên stub cha áp dụng cho mọi slide row/rank sinh
ra; row_group: set theo từng `items[i]`.

`NUM`/`multiplenumber` → luôn render donut+stack như SA-Ordinal. NUM: mỗi bin `num_quantile`/range
group = 1 slice, giữ thứ tự bin, donut hole hiện "Mean: X". multiplenumber: % slice tính từ mean
per category đã normalize (không phải % respondents), không hiện "Mean: X" (không có mean tổng).

> ⚠️ **NEVER recreate or rewrite `surveyflow/steps/appendix/generate_pptx.py`**

---

## Workflow B — User requests changes

1. Read `output/SURVEY_NAME/datatable/datatable.json`
2. Confirm change: *"Tôi sẽ [mô tả]. OK không?"*
3. Modify + save `datatable.json`
4. Detect next version: check existing `vX/` folders → increment
5. Run pipeline (table-only, no `--mcp-dir`) + same `--lang` as before
6. Present result

> *"Đã cập nhật v2. File cũ v1 vẫn còn."* — NEVER overwrite existing vX.

---

## datatable.json structure

**Field order** (giữ nhất quán khi tạo/sửa item): `type` (nếu có) → `title` → `sub_title` →
`filter` (nếu có) → `tables` → `is_appendix` (nếu có) → `banner` → `stub` → field khác
(`appendix_format`, `appendix_logo`...) chèn sau `is_appendix`.

```json
[
  {
    "title": "SURVEY_NAME - Data Table",
    "sub_title": "General",
    "tables": [
      { "sheet": "Count", "cell_content": "count",      "show_sig": false, "enabled": true },
      { "sheet": "Pct",   "cell_content": "percentage", "show_sig": false, "enabled": true, "decimal": 0 }
    ],
    "banner": [
      { "label": "Total", "filter": null },
      {
        "label": "Gender", "question": "S3",
        "groups": [
          { "label": "Male",   "value": 1 },
          { "label": "Female", "value": 2 }
        ]
      }
    ],
    "stub": [
      { "question": "S3",  "label": null, "stats": ["base", "percent"] },
      { "question": "Q36", "label": null, "stats": ["base", "percent", "mean"] }
    ]
  }
]
```

Sheet tab = `{sub_title} - {sheet}` → e.g. `"General - Count"`, `"General - Pct"`.

---

## Banner rules

- Always include `{ "label": "Total", "filter": null }` as first entry
- `value` = single code; `values` = list of codes (grouping)
- `question` field = question label from `metadata.json` (e.g. `"S3"`)
- MA questions supported as banner
- `show_total: true` trên banner entry → thêm cột Total riêng cho group đó

**Filter cấp table (per-brand/segment table)** — field **`"filter"`** cùng cấp `"banner"`/`"stub"`
trên table item, áp dụng cho MỌI banner column (kể cả Total) + mọi phép tính stub của riêng
table đó. Dùng khi tạo nhiều table riêng theo brand/segment (`"is_appendix": true`, `sub_title`
tự do, VD "Hảo Hảo"):
```json
{ "title": "SURVEY_NAME - Data Table (Base: Q2 chọn Hảo Hảo)", "sub_title": "Hảo Hảo",
  "filter": { "question": "Q2", "codes": [1] }, "tables": [...], "is_appendix": true,
  "banner": [ { "label": "Total" }, ... ], "stub": [...] }
```
Total lúc này chỉ cần viết trần `{"label": "Total"}` — KHÔNG cần lặp điều kiện lọc trên từng
banner column, KHÔNG cần `"is_total"` thủ công. `"filter"` hỗ trợ lồng nhiều cấp `and`/`or`:
```json
{ "and": [ {"question":"Q2","codes":[1]}, {"or": [{"question":"SC3","codes":[12]}, {"question":"SC3","codes":[14]}]} ] }
```
`op` mặc định `"any"` (≥1 code khớp); `"all"` chỉ dùng cho câu MA (phải chọn hết các code).

> ⚠️ Sau khi tạo, luôn kiểm tra `chart_data.json`: `"total"` của vài câu trong table đó phải
> **khác rỗng** (`{"base": N>0, ...}`) trước khi báo user xong — `total: {}` (mọi câu) nghĩa là
> filter sai `question`/`codes`, hoặc (nếu dùng cách cũ `groups`/`conditions` cho Total) thiếu
> `"is_total": true` trên đúng group đó (bug thực tế đã xảy ra — `datatable.xlsx` vẫn đủ sheet,
> chỉ appendix PPTX ra 0 slide).

Advanced: `_custom_defs` + `custom_ref`, `banner_matrix`, `levels` — thêm khi user yêu cầu.

---

## Stub rules

- **`"label"` — LUÔN `null` khi Claude tự thêm câu vào stub, KHÔNG tự viết/rút gọn nội dung vào
  field này** (khác `"title"` — được phép tự đề xuất tóm tắt). `null` → fallback sang
  `question_i18n` đầy đủ trong `metadata.json`, dùng cho header bảng xlsx + footer/Q-label slide
  (cần giữ nguyên văn để tra soát). Chỉ ghi khi **user tự tay** cho giá trị cụ thể.
- `stats`: `"base"`, `"percent"`, `"nps"`, `"mean"`, `"std"`, `"se"` — **không còn `"t2b"`/`"b2b"`**,
  2 stat này đã bỏ, thay bằng group entry `{"label":"T2B"/"B2B","codes":[..],"type":"combine"}`
  trong `choices` (tự render cùng `"percent"`, xem Step 3d mục 2)
- `"title"` (mặc định `null`) trên mỗi stub entry — tiêu đề slide appendix tuỳ chỉnh; nếu để
  `null`, `chart_data.json` tự fallback sang `title_i18n` của câu đó trong `metadata.json` (Step 3a)
- Types: `SA`, `MA`, `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`, `NUM`, `multiplenumber`
- SA với `scale_class: "Ordinal"` (metadata.json, xem Step 3c/3d) → tự động thêm
  `"mean"`/group T2B-B2B/`"nps"` tuỳ range thang đo
- `NUM` (câu số đơn): `"percent"` tự sinh 1 row/giá trị số, sort giảm dần (lớn→nhỏ). **Mặc định LUÔN
  áp dụng khi thêm câu NUM vào stub** — không cần hỏi/xác nhận trước: thêm `"mean"` vào `stats` +
  thêm `"num_quantile": 4` (pipeline tự tính 4 bin ~bằng số respondent, live mỗi lần chạy table).
  - Đổi số nhóm → sửa `num_quantile`. Muốn range cố định (VD thập kỷ tuổi) thay vì chia đều
    respondent → bỏ `num_quantile`, dùng `"choices"` tĩnh (group+hidden, preview bằng
    `scripts/group_numeric.py ... --width N`). `"choices"` luôn override `num_quantile` nếu có cả 2.
  - Không group gì → bỏ cả `"choices"` lẫn `num_quantile`.
- `multiplenumber` (số theo nhiều category, VD phân bổ chi tiêu): mỗi choice có cột số riêng —
  `"percent"` = % nhập giá trị cho category đó, `"mean"/"std"/...` tính trên giá trị của riêng category
- Auto-excluded: `FT`, `instruction`, `user-name`, `user-phone`, `reward`, `record`
- Order follows `stub` array

Advanced: `row_group`, sub-question ref (`Q17_r10`), `matrix_orientation: "horizontal"` — thêm khi user yêu cầu.

**`ranking`** (PVV xếp hạng N items) → mặc định tách thành 2 stub entries (tránh 1 slide/vị trí xếp hạng ở appendix):
```json
{ "question": "Q27_2", "stats": ["base","percent"], "ranking_mode": "rank_dist", "ranking_top_n": 3 },
{ "question": "Q27_2", "stats": ["base","percent"], "ranking_mode": "any_rank" }
```
Top 3 = Rank 1/2/3 riêng (donut, mutually exclusive). Overall (`any_rank`) = % ranked ở bất kỳ vị trí nào (không cộng 100%) → appendix tự render như MA/bar chart.

---

## Tables rules

**Default (lần đầu):** chỉ Count + Pct:
```json
"tables": [
  { "sheet": "Count", "cell_content": "count",      "show_sig": false, "enabled": true },
  { "sheet": "Pct",   "cell_content": "percentage", "show_sig": false, "enabled": true, "decimal": 0 }
]
```

**Sig test** — chỉ thêm khi user yêu cầu ("bật sig", "thêm sig test", "chạy sig"):
```json
{ "sheet": "Sig", "cell_content": "percentage", "show_sig": true, "levels": [90, 95], "method": "independent", "enabled": true }
```

`decimal`: `0` → "0%" · `1` → "0.0%" · `enabled: false` → skip sheet

---

## Common requests → datatable.json changes

| User says | Claude does |
|---|---|
| "thêm income vào banner" | Add question to `banner` |
| "bỏ Q15 khỏi stub" | Remove Q15 from `stub` |
| "thêm mean std cho Q36" | Add `"mean"`, `"std"` to Q36's stats |
| "thêm tất cả câu vào stub" | Add all codeable questions |
| "bật sig test" | Add Sig sheet to `tables` |
| "tắt sig test" | Set `"enabled": false` on Sig sheet |
| "hiện 1 chữ số thập phân" | Add `"decimal": 1` to Pct sheet |
| "thêm total cho banner group" | Add `"show_total": true` to banner entry |
| "nhóm Q13/Q14/Q17 theo brand" | Use `row_group: true` |
| "brand làm header" | Add `banner_matrix: { question: "QX" }` |
| "bảng matrix horizontal" | Add `"matrix_orientation": "horizontal"` |
| "refresh data / lấy data mới" | Re-fetch → re-run with `--force-ingestion` |
| Đính kèm file `.xlsx` có 2 sheet "Question"+"Data" | Dùng `--xlsx-input` (Step 2C) — không tự viết converter |
| "tạo PPTX / appendix" | Step 6: `--list-groups` xem table nào khớp; 1 table → chạy thẳng; ≥2 → hỏi "chạy bảng nào" + option "All"; format tự dùng `"default"` (không hỏi); nếu chưa có `appendix_logo` hỏi logo Acecook/khác/none trước; `surveyflow-pptx ...` |
| "chỉ chạy appendix cho 1-2 brand cụ thể" | Sau khi hỏi (trên), `surveyflow-pptx ... --tables idx1,idx2` (idx từ `--list-groups`) |
| "tạo appendix riêng theo từng brand/nhãn hàng" | Nhiều table item, mỗi item `"is_appendix": true` + `sub_title` tự do; lọc respondent theo brand cho cả table → field `"filter"` cấp table (xem "Filter cấp table" ở Banner rules) |
| "đổi bảng nào chạy appendix" | Hỏi lại (xem Step 6) → sửa `is_appendix` các table liên quan |
| "appendix 1 brand ra 0 slide dù xlsx đủ sheet" | Kiểm tra `chart_data.json`: `"total": {}` cho mọi câu → filter cấp table sai, hoặc thiếu `"is_total": true` trên group Total (cách cũ) — xem "Filter cấp table" |
| "appendix format general / bỏ branding" | Ghi `"appendix_format": "general"` vào table item trong `datatable.json` |
| "dùng logo khách khác / không logo Q&Me" | Ghi đường dẫn ảnh vào `"appendix_logo"`, hoặc `"none"` (chỉ Q&Me), hoặc bỏ field = Acecook mặc định |
| "rút gọn tiêu đề slide / label dài quá" | Đề xuất tóm tắt ngắn → user xác nhận → ghi vào `"title"` của stub entry |
| "chạy quality check" | Step 3b: `--run-quality`, present summary |
| "sao title_i18n toàn null" | Bình thường nếu chạy pipeline thẳng không qua Claude — chỉ Claude mới điền được (Step 3a) |
| "tóm tắt lại title_i18n" | Step 3a: đọc `question_i18n`, ghi đè `title_i18n` cho từng câu (+ sub_questions matrix) |
| "export tiếng Anh" | Re-run with `--lang en` |
| "include pending profiles" | Add `--profile-status approved,pending` |
| "đổi sang tiếng Anh" | Re-run pipeline với `--lang en` |
| "phân loại câu SA Ordinal/Nominal" | Step 3c: đọc metadata.json, tự phân loại từng câu SA **và Matrix_SA**, ghi field `scale_class`, báo summary |
| "câu này sao không tự thêm mean" | Kiểm tra `scale_class` của câu đó — chỉ Ordinal mới tự thêm `"mean"` |
| "đổi Q10 thành Nominal/Ordinal" | Sửa `scale_class` của Q10 trong metadata.json, ghi đè file |
| "sao câu Ordinal này không có T2B/NPS" | Step 3d: kiểm tra range code — chỉ 1-5 tự thêm T2B/B2B, chỉ 1-10/0-10 tự thêm NPS |
| "thang đo bị đảo (không phải tần suất), mean tính sai" | Step 3d: kiểm tra `scale_high_code`, thêm `factor` theo công thức mirror |
| "câu tần suất sao không tự có mean" | Đúng — cố ý không tự thêm, phải hỏi user trước (xem ngoại lệ Step 3d) |
| "tự nhập/ignore factor cho câu tần suất" | Nhập → dùng giá trị đó, giữ Ordinal. Ignore → sửa `scale_class` thành `"Nominal"` |
| "thêm câu ranking / gộp slide ranking" | Tách 2 stub entries: Top 3 (`ranking_mode: "rank_dist", ranking_top_n: 3`) + Overall (`ranking_mode: "any_rank"`) |

---

## FT Coding Workflow

Khi user nói: *"code câu FT / code Q26 / code open-ended"*

> ⚡ **Nhắc user chuyển Effort lên Medium trở lên trước khi bắt đầu.**
> *"Bạn vui lòng chuyển Effort → Medium (Model picker). Đã chuyển chưa?"*

### FT-1 — List câu FT

```python
import json
with open('output/SURVEY/data/metadata.json') as f:
    meta = json.load(f)
ft_qs = sorted(
    [(q['label'], q.get('question_i18n', {}).get('vi', ''))
     for q in meta['questions'].values() if q.get('answer_type') == 'FT'],
    key=lambda x: x[0])
for label, vi in ft_qs:
    print(f"  {label:10} {vi[:60]}")
```
Hỏi: *"Bạn muốn code câu nào?"*

### FT-2 — Đọc responses

```python
import json, pandas as pd
Q_LABEL = 'Q26'; SURVEY = 'VN8963'
df = pd.read_csv(f'output/{SURVEY}/data/rawdata.csv', low_memory=False)
s = df[Q_LABEL].dropna()
s = s[s.str.strip().str.len() > 2]
s = s[s.str.lower() != 'test']
responses = list(s)
with open(f'output/{SURVEY}/data/{Q_LABEL}_responses.json', 'w', encoding='utf-8') as f:
    json.dump(responses, f, ensure_ascii=False, indent=2)
print(f"{Q_LABEL}: {len(responses)} valid responses")
for i, r in enumerate(responses[:30], 1):
    print(f"  {i:3}. {r[:100]}")
```

### FT-3 — Code frame + assign + summary

```python
import re, json, pandas as pd
Q_LABEL = 'Q26'; SURVEY = 'VN8963'
with open(f'output/{SURVEY}/data/{Q_LABEL}_responses.json', encoding='utf-8') as f:
    responses = json.load(f)

CODE_FRAME = [
    (1,  "Label EN",      "Mô tả tiếng Việt"),
    (99, "Other/Unclear", "Khác/không rõ"),
]
RULES = [
    (1, r'pattern_a|pattern_b'),
]

def assign_codes(text):
    t = text.lower()
    assigned = [cid for cid, p in RULES if re.search(p, t)]
    return assigned if assigned else [99]

all_codes = [assign_codes(r) for r in responses]
n = len(responses)
counts = {}
for codes in all_codes:
    for c in codes: counts[c] = counts.get(c, 0) + 1

print(f"\n📊 FT Coding — {Q_LABEL} (n={n})\n" + "="*50)
for cid, en, vi in sorted(CODE_FRAME, key=lambda x: -counts.get(x[0], 0)):
    cnt = counts.get(cid, 0)
    print(f"  {cid:2}. {en:<28} {cnt:3} ({cnt/n*100:4.1f}%)")
```

Nếu Other/Unclear > 15% → đề xuất review thêm rules.

### FT-4 — Add vào rawdata

```python
df = pd.read_csv(f'output/{SURVEY}/data/rawdata.csv', low_memory=False)
new_cols = {f'{Q_LABEL}_coded_{cid}': pd.Series(0, index=df.index) for cid, _, _ in CODE_FRAME}
df = pd.concat([df, pd.DataFrame(new_cols)], axis=1)
q_mask = (df[Q_LABEL].notna() & (df[Q_LABEL].str.strip().str.len() > 2) & (df[Q_LABEL].str.lower() != 'test'))
for i, idx in enumerate(df[q_mask].index.tolist()):
    code_set = set(all_codes[i])
    for cid, _, _ in CODE_FRAME:
        df.at[idx, f'{Q_LABEL}_coded_{cid}'] = 1 if cid in code_set else 0
df.to_csv(f'output/{SURVEY}/data/rawdata.csv', index=False)
print(f"✅ rawdata: {len(df)} rows, {len(df.columns)} cols")
```

> ⚠️ **KHÔNG chạy `--force-ingestion` sau bước này** — sẽ mất toàn bộ cột coded.

### FT-5 — Inject vào metadata.json

```python
with open(f'output/{SURVEY}/data/metadata.json') as f: meta = json.load(f)
qs = meta['questions']
ft_pos = next((q['position'] for q in qs.values() if q.get('label') == Q_LABEL), 999)
fake_qid = f'coded_{Q_LABEL}'
qs[fake_qid] = {
    "position": ft_pos + 0.5, "question_id": 900000 + hash(Q_LABEL) % 10000,
    "label": f"{Q_LABEL}_coded",
    "question_i18n": {"vi": f"{Q_LABEL} - [Nội dung] (coded)", "en": f"{Q_LABEL} - [Text] (coded)"},
    "answer_type": "MA", "mandatory": False, "status": 1,
    "choices_i18n": {str(cid): {"vi": vi, "en": en} for cid, en, vi in CODE_FRAME},
    "rawdata_columns": [f"{Q_LABEL}_coded_{cid}" for cid, _, _ in CODE_FRAME]
}
meta['questions'] = qs
with open(f'output/{SURVEY}/data/metadata.json', 'w', encoding='utf-8') as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)
print(f"✅ metadata: rawdata_columns = {qs[fake_qid]['rawdata_columns']}")
```

> `rawdata_columns` phải khớp chính xác tên cột trong rawdata — sai tên → bảng có header nhưng không có số.

### FT-6 — Thêm stub

```python
with open(f'output/{SURVEY}/datatable/datatable.json') as f: dt = json.load(f)
dt[0]['stub'] = [s for s in dt[0]['stub'] if s.get('question') != f'{Q_LABEL}_coded']
dt[0]['stub'].append({"question": f"{Q_LABEL}_coded", "label": f"{Q_LABEL} (coded)", "stats": ["base", "percent"]})
with open(f'output/{SURVEY}/datatable/datatable.json', 'w', encoding='utf-8') as f:
    json.dump(dt, f, ensure_ascii=False, indent=2)
```

### FT-7 — Chạy bảng

```python
from surveyflow.cli import main
import os, re as _re
existing = [d for d in os.listdir(f'output/{SURVEY}') if _re.match(r'^v\d+$', d)]
next_v = f"v{max([int(d[1:]) for d in existing], default=0) + 1}"
main(['--output-dir', f'output/{SURVEY}', '--profile-status', 'approved,pending', '--version', next_v])
```

**Fix % formatting nếu cần** (nếu thấy `0.082` thay vì `8%` trong sheet Pct):
```python
import openpyxl
wb = openpyxl.load_workbook(f'output/{SURVEY}/{next_v}/datatable.xlsx')
for ws in [wb[s] for s in wb.sheetnames if 'Count' not in s]:
    in_coded = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == f'{Q_LABEL}_CODED': in_coded = True
            if in_coded and isinstance(cell.value, float) and 0 < cell.value <= 1.0:
                cell.value = round(cell.value * 100, 1); cell.number_format = '0'
wb.save(f'output/{SURVEY}/{next_v}/datatable.xlsx')
```

**FT common requests:**

| User says | Claude does |
|---|---|
| "code câu FT" | FT-1: list câu FT, hỏi user chọn |
| "code Q26" | FT-2 → FT-7 cho Q26 |
| "sửa code frame" | Update CODE_FRAME + RULES, re-run FT-4 → FT-7 |
| "review Other/Unclear" | Lấy responses code 99, in 20 examples, đề xuất rules |
| "add câu FT đã code vào bảng" | Chỉ FT-5 → FT-7 (skip FT-2/3/4) |

---

## Error handling

Khi gặp lỗi, luôn:
1. Đọc traceback, xác định dòng lỗi
2. Giải thích ngắn gọn bằng tiếng Việt
3. Đề xuất fix cụ thể

| Error | Likely cause | Fix |
|---|---|---|
| `FileNotFoundError: rawdata.csv` | Ingestion chưa chạy | Chạy lại Step 3 |
| `KeyError: 'Q5'` | Question label sai trong datatable.json | Kiểm tra `metadata.json` |
| Output có 0 rows | `profile_status` filter quá hẹp | Thêm `--profile-status approved,pending` |
| `ModuleNotFoundError: surveyflow` | Package chưa install | `pip install surveyflow` |
| `PermissionError: datatable.xlsx` | File đang mở trong Excel | Đóng Excel trước khi chạy |
| `ValueError: No data_export.csv` | Fetch chưa hoàn thành | Fetch lại hoặc upload zip |
| `json.JSONDecodeError` | datatable.json bị lỗi syntax | Đọc và kiểm tra file JSON |

---

## Confirm before acting

| Situation | Action |
|---|---|
| User gọi `/fieldcheck-dp run pipeline` | Chạy thẳng — không confirm |
| Fetch + ingestion lần đầu | Chạy thẳng — không confirm |
| User yêu cầu sửa `datatable.json` | Confirm: *"Tôi sẽ [thay đổi]. OK không?"* |
| Fetch lại data mới (overwrite) | Confirm — destructive action |
| Increment version | Thông báo: *"Tôi sẽ tạo v2 — v1 vẫn giữ nguyên"* |

- `"yes / ok / xác nhận / làm đi / chạy đi"` → proceed
- `"no / thôi / đổi lại"` → stop, hỏi user muốn đổi gì
