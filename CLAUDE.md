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
| Code FT questions | "Tôi sẽ classify [N] responses cho câu [Q_label] theo codelist [X codes]. Bạn xác nhận không?" |

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

### Step 2b — Input thay thế: file Excel "Question"+"Data" 2-sheet

Một số survey không có sẵn `definition.json` (MCP) mà chỉ có 1 file `.xlsx` gồm đúng 2
sheet **"Question"** (danh sách câu hỏi dạng phẳng — mỗi dòng là 1 câu hoặc 1 sub-item:
choice/matrix row/rank-pool item) và **"Data"** (khớp định dạng `data_export.csv` chuẩn
QMe — dòng header có ô đầu là `"Approve"`). Khi user đính kèm file dạng này, dùng thẳng
`--xlsx-input` — surveyflow tự convert sang `definition.json` + `data_export.csv` rồi chạy
ingestion bình thường, **không cần Claude tự viết converter riêng nữa**:

```bash
python run_pipeline.py \
  --xlsx-input "SURVEY_NAME.xlsx" \
  --mcp-dir    output/SURVEY_NAME/mcp \
  --output-dir output/SURVEY_NAME \
  --version    v1
```

Converter (`surveyflow/steps/ingestion/flat_xlsx_import.py`) tự suy luận SA/MA/FT/RANKING/
Matrix_SA/multiplenumber từ cấu trúc cột thật trong sheet Data (xem docstring module để biết
đầy đủ heuristic), tự loại các field interviewer-only (`[Interviewer...`/`DO NOT ask`) và
field ảnh (label kết thúc `_PHOTO`). **Luôn đọc phần `WARNING` in ra console sau khi chạy** —
đây là danh sách mọi cột bị skip/giả định (VD nhóm bị gộp thành 1 cột nhưng suy luận là
multiplenumber) để kiểm tra lại, đặc biệt với survey có cấu trúc câu hỏi khác biệt lớn so với
những gì converter đã từng gặp (mới verify với `SA`/`MA`/`FT`/`RANKING`/`Matrix_SA`/
`multiplenumber` — chưa gặp `Matrix_MA`, `NUM`, hay các synthetic type như gender/area/
personal-income).

> ⚠️ Nếu ô nào trong sheet Data có xuống dòng (Alt+Enter), converter tự thay bằng khoảng
> trắng khi dump ra CSV — cần thiết vì `export_parser.parse_export_csv` tách dòng theo
> newline vật lý, xuống dòng trong ô sẽ làm vỡ cấu trúc cột nếu không xử lý.

> Sau khi ingestion xong, **vẫn cần chạy Step 3a/3b/3c như bình thường** (title_i18n,
> scale_class, mean/factor/T2B/NPS) — converter chỉ tạo `definition.json`/`metadata.json`
> đúng structure, không tự phân loại Ordinal/Nominal hay tóm tắt tiêu đề.

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

### Step 3a — Tóm tắt tiêu đề câu hỏi (title_i18n)

Ngay sau khi có `metadata.json` (Step 3), mỗi entry câu hỏi (và mỗi entry trong `sub_questions`
của câu matrix) đã có sẵn field **`title_i18n`** do `parse_metadata()` (code, không phải AI) tự
ghi — luôn là `null` lúc mới sinh ra, vì code không có khả năng tóm tắt ngôn ngữ tự nhiên:

```json
"795699": {
  "label": "S1_1",
  "question_i18n": { "vi": "...", "en": "..." },
  "title_i18n": null,
  ...
}
```

**Claude (khi chạy qua skill, tức có AI) điền field này ngay sau ingestion** — nếu `metadata.json`
đã có `title_i18n` khác `null` cho toàn bộ câu hỏi thì bỏ qua bước này (idempotent, giống Step 3b
với `scale_class` — chỉ cần chạy **một lần** sau ingestion, không chạy lại ở các lần sửa
`datatable.json` sau này):

1. Đọc `question_i18n` (cả `vi` lẫn `en`) của từng câu trong `metadata.json`.
2. Viết 1 tiêu đề ngắn gọn (5-10 từ) tóm tắt đúng nội dung câu hỏi, giữ nguyên ý, bỏ phần kỹ
   thuật/lặp ("(Multiple answers allowed)", "SHOW ANSWER OF...", placeholder kiểu
   `{QD9/802773/Selected}`) — làm **cho cả `vi` lẫn `en`**, ghi vào đúng field đó:
   ```json
   "title_i18n": { "vi": "Lý do chọn thương hiệu", "en": "Reasons for brand choice" }
   ```
3. **Câu matrix (`sub_questions`)** — mỗi sub-question cũng có field `title_i18n` riêng (không
   có sẵn `question_i18n` như câu cha). Ghép tiêu đề của câu cha + `row_label` của dòng đó, VD
   cha có `title_i18n.en = "Brand satisfaction"`, `row_label = "Brand A"` →
   `"title_i18n": { "vi": "...", "en": "Brand satisfaction — Brand A" }`.
4. Ghi đè lại `metadata.json` với các `title_i18n` đã điền.

**Nếu chạy `surveyflow` không qua Claude** (VD chạy thẳng `run_pipeline.py`/`surveyflow-run` từ
terminal, không có AI) → `title_i18n` giữ nguyên `null` cho mọi câu, vì bước tóm tắt chỉ Claude
mới làm được — đây là hành vi đúng theo thiết kế, không phải lỗi.

> `title_i18n` **đã nối vào `chart_data.json`**: table step ghi field `"title"` (ngay sau
> `"question"`) cho mỗi câu — lấy `"title"` của stub trong `datatable.json` nếu có, không thì tự
> fallback sang `title_i18n[lang]` của câu đó trong `metadata.json` (giống hệt cách `"label"`
> fallback sang `question_i18n`). Nhờ vậy chỉ cần làm Step 3a **một lần** sau ingestion là mọi
> slide appendix sau này tự có tiêu đề tóm tắt — xem "Tiêu đề slide tuỳ chỉnh" trong Workflow D
> để biết khi nào cần ghi đè thủ công field `"title"` cấp stub.

### Step 3b — Classify SA AND Matrix_SA questions (Ordinal vs Nominal)

> ⚠️ **Quét CẢ `SA` lẫn `Matrix_SA` — đây là lỗi thực tế đã xảy ra** (một agent khác từng lọc
> cứng `answer_type == "SA"` trong script, bỏ sót toàn bộ Matrix_SA khỏi vòng phân loại, khiến
> Q17/Q15/Q11/Q6_A-kiểu-câu không bao giờ có `scale_class` → không bao giờ tự thêm mean/T2B/NPS
> ở Step 3c/3d). Hai loại này dùng **chung một quy trình phân loại**, chỉ khác chỗ ghi field.

Ngay sau khi có `metadata.json` (và `metadata.json` chưa có field `scale_class`), Claude đọc
từng câu `answer_type: "SA"` **và** `answer_type: "Matrix_SA"`, tự phân loại — dùng khả năng đọc
hiểu ngữ cảnh câu hỏi + choices, **không phải** một rule cố định trong code:

| Loại câu hỏi         | Tên nên dùng                 | Lưu vào `scale_class` |
| --------------------- | ----------------------------- | ---------------------- |
| SA có thang đo         | SA Scale / SA Likert Scale    | `"Ordinal"`             |
| SA phân loại thường    | SA Categorical / SA Nominal   | `"Nominal"`             |

Ghi giá trị vào field mới **`scale_class`** ngay trong entry của câu đó trong `metadata.json`.
**Áp dụng cho cả `SA` lẫn `Matrix_SA`** — chỉ các loại KHÁC (`MA`, `Matrix_MA`, `FT`, `NUM`,
`multiplenumber`, `ranking`...) mới bỏ qua field này:

```json
"795699": {
  "answer_type": "SA",
  "label": "S5",
  ...
  "scale_class": "Ordinal"
}
```

**Matrix_SA — ghi vào CẢ entry cha lẫn từng entry trong `sub_questions`** (thiếu phần
`sub_questions` là sai — chỉ ghi entry cha không đủ). Mỗi câu Matrix_SA có 1 bộ
`choices_i18n.columns` dùng chung cho tất cả rows (VD: thang hài lòng áp dụng cho từng brand),
nên chỉ cần phân loại **một lần** dựa trên `columns` rồi copy `scale_class` xuống từng
`sub_questions` entry — vì appendix PPTX khớp theo mã sub-question (VD `A4_r1`), không phải mã
câu cha:

```json
"802749": {
  "answer_type": "Matrix_SA", "label": "A4", "scale_class": "Ordinal",
  "sub_questions": {
    "A4_r1": { "label": "A4_r1", "scale_class": "Ordinal", ... },
    ...
  }
}
```

Sau khi thêm field cho tất cả câu SA/Matrix_SA, ghi đè lại `metadata.json`. Bước này chỉ cần
chạy **một lần** sau ingestion — nếu `scale_class` đã tồn tại trong file (từ lần chạy trước) thì
bỏ qua.

> Ví dụ Ordinal: thang hài lòng (1–5), tần suất (Never→Always), mức độ đồng ý, purchase
> intent (Definitely won't buy→Definitely will buy).
> Ví dụ Nominal: giới tính, khu vực, thương hiệu, loại sản phẩm — không có thứ tự nội tại.

**⚠️ Kiểm tra chiều thang đo — đừng giả định code lớn = điểm cao.** Một số thang đo (đặc
biệt là tần suất) đánh số NGƯỢC: code 1 = tần suất cao nhất (VD: "Hầu như mỗi ngày"), code
lớn nhất = thấp nhất (VD: "Ít hơn 1 lần/tháng"). Nếu sort mặc định theo code giảm dần sẽ ra
thứ tự SAI (thấp→cao thay vì cao→thấp). Khi thấy trường hợp này, thêm field
**`scale_high_code`** = code đại diện đầu "cao nhất" của thang đo:

```json
{ "answer_type": "Matrix_SA", "label": "A4", "scale_class": "Ordinal", "scale_high_code": 1 }
```

Bỏ qua field này nếu code lớn nhất đã là đầu "cao nhất" (trường hợp thường gặp — satisfaction,
agreement, purchase intent đều theo chiều này).

### Step 3c — Detect Ordinal (tự động thêm stats/factor/group codes)

Ngay sau Step 3b (đã có `scale_class`/`scale_high_code` trong `metadata.json`), khi thêm
một câu SA/Matrix_SA có `scale_class: "Ordinal"` vào `stub` của `datatable.json`, Claude tự
động bổ sung cấu hình sau — không cần user yêu cầu riêng:

**1. Tất cả câu Ordinal → thêm `"mean"` vào `stats` + LUÔN thêm `"factor"` cho từng choice
(bắt buộc, không có ngoại lệ).**

> ⚠️ **`mean` cho SA/Matrix_SA bắt buộc phải có `factor`** — pipeline (`table_generator.py`)
> không có đường fallback tính mean trực tiếp trên code gốc; nếu thiếu `factor` cho dù chỉ
> 1 choice, toàn bộ `mean` của câu đó sẽ trả về **0.0** một cách âm thầm (không lỗi). Vì vậy
> **luôn** thêm `factor` cho mọi code thật của thang đo (bỏ qua meta code 98/99), kể cả khi
> thang đo không đảo ngược.
>
> **Chỉ duy nhất 1 format được hỗ trợ:** per-choice `"choices": [{"code": .., "factor": ..}]`
> như ví dụ dưới đây. Pipeline **không** nhận bất kỳ field nào khác cho mục đích này — không
> có top-level `"factors"` (dict), không có `"mean_factor"` (dict), không có `"factor"` số ít
> ở top-level. Đây là bug thực tế đã xảy ra 2 lần (dùng nhầm `"factor"` số ít, và trước đó
> pipeline từng có 2 field fallback này nhưng đã bị bỏ hẳn) — đều dẫn tới `mean` âm thầm = 0.0.
>
> **Matrix_SA — lấy code từ `choices_i18n.columns`, KHÔNG BAO GIỜ từ `.rows`.** `.rows` là
> danh sách item được đánh giá (VD brand) — không phải thang đo, và không dùng để tính
> `factor`/`scale_class`/`t2b_codes`/NPS groups. Đây cũng là bug thực tế đã xảy ra (nhầm lấy
> code từ `.rows` khiến factor/scale hoàn toàn sai vì số lượng item ≠ số điểm thang đo).

> ⚠️ **NGOẠI LỆ — câu tần suất (frequency scale): KHÔNG tự động thêm mean/factor.** Thang tần
> suất (VD "Hàng ngày / Vài lần tuần / Hiếm khi / Không bao giờ", "Never→Always") thường
> không có 1 công thức factor "đúng" hiển nhiên — khoảng cách giữa các mức không đều nhau về
> mặt thời gian thực tế, hoặc có mức bất thường phá vỡ thứ tự thuần tuý (VD "Bị trả chậm" xen
> giữa thang tần suất — xem ví dụ ở Step 3b). Gán `factor` tự động (identity/mirror) cho loại
> câu này có thể ra mean sai lệch về mặt ý nghĩa dù công thức tính đúng cú pháp. Khi Step 3b
> phân loại 1 câu SA/Matrix_SA là Ordinal **và nội dung là thang tần suất**:
> 1. **Không** áp dụng mục 1-3 dưới đây tự động như các câu Ordinal khác.
> 2. Hỏi user: *"Câu {label} là thang đo tần suất. Bạn muốn tự nhập factor cho từng mức (VD:
>    Hàng ngày=5, Vài lần/tuần=4, Hiếm khi=2, Không bao giờ=1) hay để pipeline coi câu này là
>    Nominal (không tính mean)?"*
> 3. User cung cấp factor → dùng đúng giá trị user đưa (per-choice, cùng format `choices` như
>    trên), giữ `scale_class: "Ordinal"` trong `metadata.json`, tiếp tục áp dụng T2B/B2B/NPS
>    bình thường theo range (mục 2-3 dưới đây).
> 4. User chọn bỏ qua ("ignore"/"không cần") → sửa `scale_class` trong `metadata.json` từ
>    `"Ordinal"` thành `"Nominal"` cho câu đó — SA sửa entry đó; Matrix_SA sửa **cả entry cha
>    lẫn từng `sub_questions`** (giống quy tắc ghi ở Step 3b). Stub chỉ giữ
>    `["base", "percent"]`, không thêm `mean`/`factor`/T2B/B2B/NPS.

Thang đo **không** đảo ngược (`scale_high_code` bỏ trống hoặc = code lớn nhất) → `factor`
= chính code đó (identity mapping):

```json
{ "question": "C4", "stats": ["base", "percent", "mean"],
  "choices": [
    { "code": 1, "factor": 1 },
    { "code": 2, "factor": 2 },
    { "code": 3, "factor": 3 },
    { "code": 4, "factor": 4 },
    { "code": 5, "factor": 5 }
  ]
}
```

Thang đo **bị đảo ngược** (`scale_high_code` = code nhỏ nhất, không phải lớn nhất trong các
code thật của thang đo) → `factor` mirror để mean phản ánh đúng chiều (factor cao = giá trị cao).
(Ví dụ dưới đây là thang đồng ý 7 điểm bị đảo — **không** phải câu tần suất, xem ngoại lệ ở trên
cho câu tần suất):

```
factor(code) = code_max + code_min - code
```

```json
{ "question": "A4", "label": "Agreement (reversed)", "stats": ["base", "percent", "mean"],
  "choices": [
    { "code": 1, "factor": 7 },
    { "code": 2, "factor": 6 },
    { "code": 3, "factor": 5 },
    { "code": 4, "factor": 4 },
    { "code": 5, "factor": 3 },
    { "code": 6, "factor": 2 },
    { "code": 7, "factor": 1 }
  ]
}
```

**2. Thang đo 1–5 → thêm T2B/B2B:**

> ⚠️ **T2B/B2B KHÔNG còn là stat riêng** (`"t2b"`/`"b2b"` trong `stats`) và KHÔNG còn dùng field
> `t2b_codes`/`b2b_codes` — đây là format cũ, pipeline hiện **bỏ qua lặng lẽ** nếu còn gặp (không
> lỗi, chỉ đơn giản không sinh hàng T2B/B2B nữa). T2B/B2B giờ là 1 **group entry bình thường**
> ngay trong `choices` — dùng chung cơ chế group tổng quát (giống group "Increased"/"Decreased"
> tuỳ ý ở Step 3d mục khác, hay group NUM range ở Stub rules). `stats` chỉ cần `"percent"` — group
> luôn tự render kèm theo, không cần khai báo `"t2b"`/`"b2b"` trong `stats`.

2 code gần đầu "cao nhất" của thang (theo `scale_high_code`) → group `"label": "T2B"`; 2 code gần
đầu "thấp nhất" → group `"label": "B2B"`. Thêm 2 entry này vào **cuối** mảng `choices` (sau các
entry `{"code":..,"factor":..}` per-choice):

```json
{ "question": "C4", "stats": ["base", "percent", "mean"],
  "choices": [
    { "code": 1, "factor": 1 },
    { "code": 2, "factor": 2 },
    { "code": 3, "factor": 3 },
    { "code": 4, "factor": 4 },
    { "code": 5, "factor": 5 },
    { "label": "T2B", "codes": [4, 5], "type": "combine" },
    { "label": "B2B", "codes": [1, 2], "type": "combine" }
  ]
}
```

Đảo ngược (`scale_high_code: 1`) → đảo luôn:
`{ "label": "T2B", "codes": [1, 2], "type": "combine" }`,
`{ "label": "B2B", "codes": [4, 5], "type": "combine" }`.

**3. Thang đo 1–10 hoặc 0–10 → thêm NPS groups + stat `"nps"`:**

2 code gần đầu "cao nhất" → Promoters; 2 code kế tiếp → Passives; các code còn lại (đầu
"thấp nhất") → Detractors. Stat `"nps"` **vẫn cần khai báo riêng** trong `stats` (khác T2B/B2B —
NPS là 1 phép tính điểm số `%Promoters − %Detractors`, không phải chỉ 1 hàng % gộp đơn thuần, nên
không tự động chạy chỉ nhờ có group trong `choices`):

```json
{ "question": "E9", "stats": ["base", "percent", "mean", "nps"],
  "choices": [
    { "type": "promoters",  "codes": [9, 10],              "label": "Promoters" },
    { "type": "passive",    "codes": [7, 8],                "label": "Passives" },
    { "type": "detractors", "codes": [0, 1, 2, 3, 4, 5, 6], "label": "Detractors" }
  ]
}
```

(Thang chỉ có 1–10, không có code 0 → Detractors = `[1,2,3,4,5,6]`.) Đảo ngược → Promoters
luôn là 2 code gần đầu "cao nhất" theo `scale_high_code` (có thể là code nhỏ), Detractors là
phần code ở đầu "thấp nhất" — không tính theo trị số tuyệt đối.

**Giá trị hợp lệ cho `"type"`** trong 1 group entry (mảng `choices`, phân biệt hoa/thường, luôn
viết thường): `"combine"` (mặc định, gộp thành 1 hàng % duy nhất — dùng cho T2B/B2B và mọi group
tuỳ ý khác), `"netted"` (gộp thành 1 hàng % TỔNG, cộng thêm các hàng % riêng từng code bên dưới),
`"promoters"`/`"detractors"` (chỉ 2 giá trị này được stat `"nps"` đọc để tính điểm số — xem mục 3).
`"passive"` (hoặc bất kỳ chuỗi nào khác) không có ý nghĩa đặc biệt với code — chỉ hiển thị như 1
group `"combine"` bình thường, dùng cho mục đích trình bày (VD tách riêng nhóm Passives ở giữa).

**Câu Ordinal có thang đo khác 1–5 và 1–10/0–10** → chỉ thêm `"mean"` (mục 1), không tự
thêm T2B/B2B/NPS.

**Luôn bỏ qua meta code 98/99** (Others/No answer) khỏi mọi phép tính trên — chúng không
thuộc thang đo.

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

→ Nếu "all": thêm tất cả SA/MA/Matrix/NUM/multiplenumber theo thứ tự position, stats mặc định
  `["base", "percent"]`
→ Nếu chỉ định cụ thể: chỉ thêm các câu đó theo đúng thứ tự user nhập
→ **Câu SA/Matrix_SA có `scale_class: "Ordinal"` trong metadata.json** → tự động thêm `"mean"` vào
  `stats` (`["base", "percent", "mean"]`), không cần user yêu cầu riêng. Câu Nominal hoặc
  không có `scale_class` thì giữ mặc định `["base", "percent"]`.
→ **Câu `NUM`** → LUÔN tự động thêm `"mean"` vào `stats` + thêm `"num_quantile": 4` vào stub entry,
  không cần user yêu cầu riêng (xem Stub rules mục NUM).
→ **Câu `multiplenumber`** → LUÔN tự động thêm `"mean"` vào `stats` (`["base", "percent", "mean"]`),
  không cần user yêu cầu riêng (xem Stub rules mục multiplenumber).

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

## Workflow C — FT (Open-ended) Analysis

When user says: *"code câu FT"*, *"analyze open-ended"*, *"tạo codelist"*, *"code câu mở Q5"*, *"classify responses"*

### Step 1 — Identify FT questions

Read `output/SURVEY_NAME/data/metadata.json` → find all questions with type `FT`.

Display list and ask:
> "Tìm thấy X câu FT: Q5 (Lý do chọn), Q10 (Khác - ghi rõ). Bạn muốn code tất cả hay chỉ một số câu?"

### Step 2 — Create codelist (per FT question)

**Nếu user cung cấp codelist** → dùng đúng code numbers của user, bỏ qua bước generate.

**Nếu user không cung cấp → Thematic Analysis (recommended):**
- Đọc **tất cả** responses của câu FT từ `rawdata.csv` (không chỉ sample)
- Claude phân tích toàn bộ → phát hiện themes bottom-up → đề xuất codelist
- Ưu điểm so với sampling: không bỏ sót minority themes, codelist gắn với data thực tế
- User review/chỉnh sửa → user xác nhận trước khi classify

Luôn thêm 2 code cuối:
- `{ "code": 98, "label": "Others" }` — response không khớp code nào
- `{ "code": 99, "label": "No answer" }` — blank / NA / không trả lời

Save mỗi câu một file: `output/SURVEY_NAME/data/ft_codelist_{Q_label}.json`

```json
{
  "question": "Q5",
  "label": "Lý do chọn sản phẩm",
  "codes": [
    { "code": 1, "label": "Giá cả / Tiết kiệm" },
    { "code": 2, "label": "Chất lượng tốt" },
    { "code": 3, "label": "Tin tưởng thương hiệu" },
    { "code": 98, "label": "Others" },
    { "code": 99, "label": "No answer" }
  ]
}
```

### Step 3 — Classify (batch tất cả FT questions)

Xử lý lần lượt từng FT question. Với mỗi câu:
- Đọc cột FT từ `rawdata.csv`
- Batch 20–30 responses mỗi lần
- Multi-code: 1 response có thể nhận nhiều code
- Code 99 nếu response rỗng / null / "N/A" / không trả lời
- Code 98 nếu response có nội dung nhưng không khớp code nào

**Prompt pattern cho classification:**
```
Codelist:
1 = [label]
2 = [label]
98 = Others (có nội dung nhưng không khớp code nào)
99 = No answer (trống / NA / không trả lời)

Classify từng response bên dưới. Multi-code được phép.
Trả về JSON: [{"id": "R001", "codes": [1, 3]}, ...]

Responses:
R001: "[text]"
R002: "[text]"
...
```

### Step 4 — Output

Write `output/SURVEY_NAME/data/ft_coded.csv`:
- Cột `resp_id` match với `rawdata.csv`
- Binary columns: `{Q_label}_c{code}` → giá trị `1` hoặc `0`
- Tất cả FT questions trong cùng 1 file

```
resp_id, Q5_c1, Q5_c2, Q5_c3, Q5_c98, Q5_c99, Q10_c1, Q10_c98, Q10_c99
R001,    1,     0,     1,     0,      0,      1,      0,       0
R002,    0,     1,     0,     0,      0,      0,      1,       0
R003,    0,     0,     0,     0,      1,      0,      0,       1
```

> **Sau khi xong:** Hỏi user có muốn thêm các cột FT coded vào datatable không.
> Nếu có → đợi user yêu cầu cụ thể (Workflow B).

---

## datatable.json structure

`datatable.json` is an **array** — mỗi item là 1 table config độc lập, sinh ra các sheets riêng trong cùng 1 file xlsx.

**Kiểu item trong array:**
- `{ "type": "datatable", ... }` → bảng chéo thông thường (hoặc bỏ qua `type`, default là datatable)
- `{ "_custom_defs": [...] }` → khối định nghĩa filter dùng chung (không sinh sheet, chỉ dùng để tham chiếu)

**Thứ tự field chuẩn trong 1 table item** — giữ nhất quán khi Claude tạo mới hoặc sửa
`datatable.json` (không bắt buộc với field không có trong item, và `type` thường bỏ qua vì
default đã là `"datatable"`):
`type` (nếu có) → `title` → `sub_title` → `filter` (nếu có) → `tables` → `is_appendix` (nếu có)
→ `banner` → `stub` → các field khác (`appendix_format`, `appendix_logo`, `matrix_orientation`,
`matrix_rows`, `sig_direction`, `banner_matrix`...) chèn ngay sau `is_appendix`.

```json
[
  {
    "title": "SURVEY_NAME - Data Table",
    "sub_title": "General",
    "tables": [
      { "sheet": "Count", "cell_content": "count",      "show_sig": false, "enabled": true },
      { "sheet": "Pct",   "cell_content": "percentage", "show_sig": false, "enabled": true, "decimal": 0 },
      { "sheet": "Sig",   "cell_content": "percentage", "show_sig": true,  "levels": [90, 95], "method": "independent", "enabled": true }
    ],
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
      { "question": "Q10", "label": null, "title": null, "stats": ["base", "percent"] },
      { "question": "Q36", "label": null, "title": null, "stats": ["base", "percent", "mean"],
        "choices": [
          { "code": 1, "factor": 1 }, { "code": 2, "factor": 2 }, { "code": 3, "factor": 3 },
          { "code": 4, "factor": 4 }, { "code": 5, "factor": 5 },
          { "label": "T2B", "codes": [4, 5], "type": "combine" },
          { "label": "B2B", "codes": [1, 2], "type": "combine" }
        ]
      }
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

#### Filter cấp table (per-brand/segment table) — field `"filter"` trên table item

Bảng bình thường (`{ "label": "Total", "filter": null }`, không hậu tố, không cần lọc gì) →
không cần đọc mục này.

Khi tạo **nhiều table riêng theo từng brand/segment** (VD Workflow D — mỗi brand/segment 1 table
với `"is_appendix": true`, xem bảng common-requests), **toàn bộ respondent của table đó** cần lọc
theo 1 điều kiện chung (VD chỉ tính người dùng brand X, `Q2 == mã brand`). Dùng field **`"filter"`**
ngay cấp table (cùng cấp `"banner"`/`"stub"`) — áp dụng cho **mọi banner column** (kể cả Total)
**và** mọi phép tính stub (base/percent/mean) của riêng table đó, các table khác trong cùng
`datatable.json` không bị ảnh hưởng:

```json
{
  "title": "SURVEY_NAME - Data Table (Base: Q2 chọn Hảo Hảo)",
  "sub_title": "Hảo Hảo",
  "filter": { "question": "Q2", "codes": [1] },
  "tables": [...],
  "is_appendix": true,
  "banner": [
    { "label": "Total" },
    { "label": "Khu vực", "question": "SC3", "groups": [...] }
  ],
  "stub": [...]
}
```

Nhờ filter áp dụng trước khi tính banner, cột Total ở đây **chỉ cần viết trần**
`{ "label": "Total" }` như bảng bình thường — KHÔNG cần lặp lại điều kiện lọc trên từng banner
column, KHÔNG cần field `"is_total"` thủ công (đơn giản hơn hẳn cách cũ ở dưới).

**Cú pháp `"filter"`** — dùng chung với `_custom_defs`/`custom_ref` (xem mục dưới), hỗ trợ lồng
nhiều cấp `and`/`or` tuỳ ý:
```json
{ "question": "Q2", "codes": [1, 2], "op": "any" }
{ "and": [ { "question": "Q2", "codes": [1] }, { "question": "SC3", "codes": [12,41,20,38] } ] }
{ "or":  [ { "question": "Q2", "codes": [1] }, { "question": "Q2", "codes": [5] } ] }
{ "and": [ { "question": "Q2", "codes": [1] },
           { "or": [ { "question": "SC3", "codes": [12] }, { "question": "SC3", "codes": [14] } ] } ] }
```
`op` mặc định `"any"` (chọn ít nhất 1 code); `"all"` chỉ có ý nghĩa với câu MA (phải chọn **hết**
các code liệt kê).

**Kiểm tra sau khi tạo**: mở `chart_data.json`, xác nhận `"total"` của vài câu trong table đó
**không rỗng** (`{"base": N>0, "percents": {...}}`) trước khi báo user là xong.

**Cách cũ (không dùng field `"filter"`)** — nếu vì lý do nào đó chỉ cần lọc RIÊNG cột Total (không
lọc cả table), vẫn có thể viết Total dạng `groups`/`conditions` như 1 breakdown bình thường +
thêm **`"is_total": true`** trên đúng group đó:
```json
{ "label": "Total", "groups": [
  { "label": "Total", "conditions": [{ "question": "Q2", "value": 1 }], "is_total": true }
]}
```
> ⚠️ Thiếu `"is_total": true` ở cách này → appendix PPTX âm thầm sinh 0 slide cho table đó dù
> `datatable.xlsx` vẫn đủ sheet (bug thực tế đã xảy ra) — vì group dạng `groups`/`conditions`
> mặc định `is_total: false`. Ưu tiên dùng `"filter"` cấp table ở trên — đơn giản hơn và áp dụng
> nhất quán cho cả banner lẫn stub, tránh quên field này.

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
- `row_code` / `row_codes` chấp nhận **số nguyên hoặc string** (pipeline tự convert sang string)
- `row_code` (single): paired mode reads `{q_col}_r{code}` per column
- `row_codes` (list): stacked mode sums counts across all `{q_col}_r{rc}` columns
- Mix of both is allowed in the same `groups` array

**Header levels produced:**
- Total column + each brand: `Total / Brand` (2 levels)
- Store Type column + each brand: `Store Type / Sub-group / Brand` (3 levels)

### matrix_orientation — Matrix rows as banner sub-columns (table-level)

Dùng `"matrix_orientation": "horizontal"` khi toàn bộ stub là các câu matrix có **cùng rows** và muốn:
- **Rows → banner sub-columns** (xuất hiện dưới mỗi banner group: Total/Brand A, Total/Brand B...)
- **Choices → stub rows** (Very satisfied, Satisfied, Neutral...)

```json
{
  "type": "datatable",
  "sub_title": "Brand Analysis",
  "matrix_orientation": "horizontal",
  "banner": [
    { "label": "Total", "filter": null },
    {
      "label": "Gender",
      "question": "S1",
      "groups": [
        { "label": "Male",   "value": 1 },
        { "label": "Female", "value": 2 }
      ]
    }
  ],
  "stub": [
    { "question": "Q17", "label": "Brand Satisfaction", "stats": ["base", "percent"] },
    { "question": "Q18", "label": "Brand Usage Freq",   "stats": ["base", "percent"] },
    { "question": "Q19", "label": "Brand Imagery",      "stats": ["base", "percent"] }
  ],
  "tables": [...]
}
```

**Output layout:**
```
                       | Total              | Male               | Female
                       | Br.A  Br.B  Br.C  | Br.A  Br.B  Br.C  | Br.A  Br.B  Br.C
Q17 Brand Satisfaction
  Very satisfied        | 30%   25%   20%   | 35%   28%   22%   | ...
  Satisfied             | 40%   38%   35%   | 38%   36%   33%   | ...
  Neutral               | 20%   25%   30%   | ...
Q18 Brand Usage Freq
  Every day             | 15%   12%    8%   | ...
  Weekly                | 30%   28%   25%   | ...
```

**Rules:**
- Tất cả stub entries phải là matrix questions (`Matrix_SA`, `Matrix_MA`, `Matrix_NUM`)
- Tất cả phải có cùng `choices_i18n.rows` (pipeline báo lỗi nếu không match)
- Không mix SA/MA thường vào stub khi dùng `matrix_orientation: "horizontal"`
- Nếu cần chạy cả câu thường lẫn matrix horizontal → tạo 2 table items riêng trong array
- Default (bỏ qua hoặc `"vertical"`) → behavior hiện tại (expand theo rows)

**So sánh với `banner_matrix`:**

| | `banner_matrix` | `matrix_orientation: "horizontal"` |
|---|---|---|
| Config | Cấp table, 1 câu matrix làm header | Cấp table, áp dụng cho tất cả stub entries |
| Rows | Từ 1 câu matrix chỉ định | Lấy từ tất cả câu trong stub (phải đồng nhất) |
| Stub | Paired mode — câu matrix trong stub dùng rows từ banner_matrix | Choices của mỗi câu thành stub rows |
| Dùng khi | Cần chọn brands cụ thể, group brands | Tất cả câu đều là matrix, cùng rows |

### matrix_rows — show/hide/combine brand sub-columns

Dùng trong table item có `matrix_orientation: "horizontal"` để chọn rows nào xuất hiện làm sub-columns, và nhóm nhiều rows thành 1 column.

```json
"matrix_rows": [
  { "row_code": 1,                        "label": "STING" },
  { "row_code": 2,                        "label": "RED BULL" },
  { "row_code": 3,                        "label": "NUMBER 1" },
  { "row_codes": [4, 5, 6, 7, 9, 10],    "label": "Others" }
]
```

**Rules:**
- `row_code` / `row_codes` chấp nhận số nguyên hoặc string
- Rows không có trong `matrix_rows` → tự động ẩn
- Bỏ qua `matrix_rows` (hoặc để `null`) → tất cả rows từ metadata đều hiện
- `row_codes` (list) → stacked mode: cộng dồn counts của tất cả rows được liệt kê

### sig_direction — chiều so sánh sig test

Dùng trong table item có `matrix_orientation: "horizontal"`.

```json
"sig_direction": "rows"
```

- `"rows"` (default): so sánh **brands với nhau** trong cùng 1 demographic group → "trong Total, STING vs RED BULL"
- `"columns"`: so sánh **demographics với nhau** trong cùng 1 brand → "với STING, Total vs Male vs Female"

> Use case `"columns"` ít gặp. Default `"rows"` là đúng cho hầu hết trường hợp.

### Stub rules
- One entry per question
- **`"label"` — LUÔN để `null` khi Claude tự thêm 1 câu vào stub, KHÔNG tự ý viết/rút gọn nội dung
  vào field này** (kể cả khi label gốc dài) — khác hẳn `"title"` (xem bên dưới), vốn được PHÉP tự
  đề xuất bản tóm tắt. `"label": null` fallback sang `question_i18n` đầy đủ trong `metadata.json`
  — dùng làm header hàng bảng xlsx **và** footer/Q-label cuối slide appendix, cần giữ nguyên văn
  để tra soát với câu hỏi khảo sát gốc. Chỉ ghi `"label"` khi **user tự tay** cung cấp giá trị cụ
  thể (VD user nói "đặt label câu S3 là 'Gender'").
- `stats` options: `"base"`, `"percent"`, `"mean"`, `"std"`, `"se"`, `"nps"` — **không còn `"t2b"`/`"b2b"`**;
  T2B/B2B giờ là group entry trong `choices` (`"type": "combine"`), tự render cùng `"percent"`, xem
  Step 3c mục 2
- `"title"` (optional, mặc định `null`) — tiêu đề ngắn gọn hiển thị trên đầu slide appendix
  (General lẫn default), thay cho label gốc dài; xem "Tiêu đề slide tuỳ chỉnh" trong Workflow D
- Supported answer types: `SA`, `MA`, `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`, `NUM`, `multiplenumber`
  - Matrix questions automatically expand into one block per row (sub-question)
  - When `banner_matrix` is active, matrix questions use paired mode instead
  - When `matrix_orientation: "horizontal"` is active, matrix rows → banner sub-columns, choices → stub rows
  - **`NUM`** (câu số đơn, VD tuổi): `"percent"` tự sinh 1 row cho từng giá trị số xuất hiện trong data,
    sort **giảm dần** (lớn → nhỏ) — khác FT (sort chữ cái, dùng cho text). `"mean"/"std"/"se"/"min"/"max"`
    tính trực tiếp trên giá trị số gốc (không cần `factor`, khác SA/Matrix_SA Ordinal).
    - **Mặc định LUÔN áp dụng khi Claude thêm câu NUM vào stub — không cần user yêu cầu riêng**
      (giống Step 3c cho Ordinal, ranking Top3+Overall): thêm `"mean"` vào `stats` + thêm
      `"num_quantile": 4` vào stub entry. Pipeline tự tính 4 bin ~bằng nhau về số respondent từ
      phân bố tổng **mỗi lần chạy table** — không cần group thủ công, không cần user xác nhận trước:
      ```json
      { "question": "S3_1", "stats": ["base", "percent", "mean"], "num_quantile": 4 }
      ```
    - **Đổi số nhóm** (VD "chia 6 nhóm"): sửa giá trị `num_quantile`.
    - **Group range cố định thay vì quantile** (VD user muốn đúng thập kỷ tuổi "20-29"/"30-39" thay vì
      chia đều respondent): bỏ `num_quantile`, thay bằng `"choices"` tĩnh theo cơ chế group thống nhất
      (giống group T2B/B2B) — mỗi giá trị gộp vào 1 group cần thêm `{ "code": "<value>", "hidden": true }`:
      ```json
      { "question": "S3_1", "stats": ["base", "percent", "mean"],
        "choices": [
          { "codes": ["20","21","22","23","24"], "label": "20-24" },
          { "codes": ["25","26","27","28","29","30"], "label": "25-30" },
          { "code": "20", "hidden": true }, { "code": "21", "hidden": true }
        ]
      }
      ```
      Dùng `python .skill_work/fieldcheck-dp/scripts/group_numeric.py <rawdata.csv> <metadata.json>
      --question S3_1 [--width N | --quantile N]` để xem trước bin/size trước khi chọn. Nếu stub
      entry có cả `"choices"` lẫn `num_quantile`, `"choices"` luôn thắng (override).
    - **Không group gì cả** (user yêu cầu giữ từng giá trị riêng lẻ): bỏ cả `"choices"` lẫn
      `num_quantile` — pipeline hiện đủ mỗi giá trị số 1 row (sort giảm dần).
  - **`multiplenumber`** (câu số theo nhiều category, VD phân bổ chi tiêu theo từng khoản mục): mỗi
    choice có 1 cột số riêng trong rawdata — `"percent"` = % respondents có nhập giá trị cho category đó;
    `"mean"/"std"/"se"/"min"/"max"` tính trên giá trị số của riêng category đó (không phải toàn câu).
- Excluded automatically: `FT`, `instruction`, `user-name`, `user-phone`, `reward`, `record`
  (FT vẫn xử lý riêng qua Workflow C — quá nhiều giá trị unique để liệt kê thẳng như NUM)
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

#### ranking — mặc định Top 3 + Overall (any rank)

Câu `answer_type: "ranking"` (PVV xếp hạng N items) **mặc định luôn tách thành 2 stub entries** khi thêm vào datatable.json — không cần user yêu cầu riêng, tự động để tránh việc appendix sinh ra 1 slide/vị trí xếp hạng (VD 16 items → 16 vị trí → rất nhiều slide, phần lớn base quá nhỏ ở các rank cao):

```json
{ "question": "Q27_2", "label": null, "stats": ["base", "percent"],
  "ranking_mode": "rank_dist", "ranking_top_n": 3 },
{ "question": "Q27_2", "label": null, "stats": ["base", "percent"],
  "ranking_mode": "any_rank" }
```

- **Top 3** (`ranking_mode: "rank_dist", ranking_top_n: 3`) → chỉ sinh block Rank 1/2/3 (bỏ Rank 4+ vì base thường quá nhỏ để báo cáo % đáng tin cậy) — mỗi rank vẫn là 1 câu SA-style donut+stack riêng (mutually exclusive: mỗi respondent có đúng 1 item ở mỗi rank).
- **Overall** (`ranking_mode: "any_rank"`) → gộp tất cả vị trí thành 1 block duy nhất, mỗi item = % respondents từng xếp hạng item đó ở BẤT KỲ vị trí nào (không mutually exclusive, giống MA) — appendix tự động render block này như **câu MA / bar chart** (không phải donut), vì % các item không cộng lại 100%.
- Muốn giữ nhiều hơn Top 3 (VD Top 5) → đổi `ranking_top_n: 5`.
- Slide title/Q-label của appendix tự ghép mã + nội dung câu ranking gốc vào từng sub-block, VD `Q27_2-Rank 1`, `Q27_2-Overall` — không cần cấu hình thêm.

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
- `SA`, `MA`, `Matrix_SA`, `Matrix_MA`, `Matrix_NUM`, `NUM`, `multiplenumber`
- Default stats: `["base", "percent"]`
- SA/Matrix_SA với `scale_class: "Ordinal"` → `["base", "percent", "mean"]`
- `NUM`/`multiplenumber` → `["base", "percent", "mean"]` (percent = category breakdown, xem Stub rules)
- `NUM` → luôn thêm `"num_quantile": 4` (xem Stub rules — quantile group mặc định, không cần hỏi)
- Skip: `FT`, `instruction`, `user-name`, `user-phone`, `reward`, `record`

---

## Workflow D — PPTX Chart Appendix (slides.pptx)

Sau khi chạy bảng (Step 5 / Workflow B), table step **tự sinh** `chart_data.json`
cạnh `datatable.xlsx`. Từ đó tạo bộ slide biểu đồ (**appendix**) — chart PowerPoint
**editable**, mỗi câu 1 slide, style khớp template công ty.

**Chọn table để chạy appendix** — appendix render **mọi table** trong `datatable.json` có field
**`"is_appendix": true`**, gộp lại thành **1 file .pptx duy nhất**. Mỗi table đánh dấu là **1
group** riêng, tên group = `sub_title` của table đó **nguyên văn** (KHÔNG còn quy ước prefix
"Appendix -"/"General" — `sub_title` giờ là tên tự do, đặt gì cũng được, VD `"Total"`, `"Hảo
Hảo"`, `"Omachi"`):

```json
{ "type": "datatable", "sub_title": "Hảo Hảo", "is_appendix": true, "banner": [...], ... }
```

- **Chỉ 1 table `is_appendix: true`** (trường hợp thường gặp) → render y như cũ, không chia
  group, không hỏi gì thêm.
- **≥2 table `is_appendix: true`** (VD nhiều brand) → mỗi group được bọc bởi 1 **PowerPoint
  Section** (xem trong Slide Sorter/Outline view của PowerPoint) đúng các slide của group đó —
  cho phép nhảy nhanh giữa các brand khi review. **Không** chèn slide riêng nào cho tên group —
  group chỉ tốn đúng số slide câu hỏi của nó, không hơn. Style (`appendix_format`)/logo
  (`appendix_logo`) áp dụng cho **cả deck** — lấy từ table `is_appendix: true` **đầu tiên** (hoặc
  override `--format`/`--logo`); nếu 1 table khác trong deck khai `appendix_format` khác thì bị
  bỏ qua kèm cảnh báo (in ra console, không lỗi).
- **Không table nào được đánh dấu** → **không chạy**, hỏi user muốn chạy bảng nào (xem mục dưới)
  rồi mới đánh dấu + chạy.
- `--table N` (CLI) hoặc `table_idx=` (API) vẫn giữ hành vi cũ: chỉ render **đúng 1 table** theo
  `table_index`, bỏ qua hoàn toàn cơ chế `is_appendix`/group ở trên (override 1 lần, không sửa
  file).

**Hỏi user muốn chạy bảng nào — LƯU câu trả lời vào `is_appendix`** (khác câu hỏi format/logo ở
chỗ: đây **không phải** hỏi mỗi lần chạy — hỏi **một lần**, lưu vào `datatable.json`, các lần
chạy appendix sau tự dùng lại y hệt format/logo, không hỏi lại trừ khi user muốn đổi bảng):

1. Liệt kê **toàn bộ** table (không chỉ table đã đánh dấu): `surveyflow-pptx <chart_data.json>
   --list-groups` → in JSON `[{"table_index", "sub_title", "is_appendix"}, ...]` cho MỌI table.
2. Hỏi user (liệt kê đúng `sub_title` đã liệt kê ở bước 1, không tự bịa tên):
   > *"Bạn muốn chạy appendix cho bảng nào? {sub_title1}, {sub_title2}, {sub_title3}...
   > (chọn 1 hay nhiều tên, hoặc 'All' để chạy tất cả)"*
3. Theo câu trả lời, **sửa `datatable.json`** — confirm với user trước khi sửa (theo quy tắc
   chung "confirm trước khi sửa datatable.json"):
   - **"All"** → set `"is_appendix": true` cho mọi table trong file (trừ `_custom_defs`).
   - **1 hoặc vài tên cụ thể** → set `"is_appendix": true` cho đúng các table được chọn; table
     nào **trước đó** đã `is_appendix: true` nhưng lần này KHÔNG được chọn lại → set `false`
     (hoặc xoá field) để không còn tự chạy nữa.
4. Sau khi đã lưu, chạy `surveyflow-pptx`/pipeline `--appendix` bình thường — **không cần**
   `--tables` nữa cho flow chuẩn (field `is_appendix` đã quyết định sẵn). `--tables idx1,idx2`
   vẫn dùng được để override **1 lần** mà không sửa file (VD test nhanh 1 table cụ thể).

### Format appendix — "default" (mặc định) hay "general"

Appendix **luôn dùng format `"default"`** (style công ty, xem "Format 'default' khác 'general'
ở đâu" bên dưới) trừ khi user chỉ định rõ dùng `"general"` (style chart mặc định thuần của
surveyflow, không có branding) — **không cần hỏi user chọn format nữa**, không như logo bên
dưới. `"general"` vẫn được giữ nguyên trong code (không xoá), chỉ không còn là lựa chọn mặc định.

Chỉ ghi field **`appendix_format`** vào table item trong `datatable.json` khi user yêu cầu đổi
khỏi mặc định (VD "chạy appendix format general"):

```json
{
  "title": "SURVEY_NAME - Data Table",
  "sub_title": "Appendix",
  "tables": [...],
  "is_appendix": true,
  "appendix_format": "general",
  "banner": [...],
  "stub": [...]
}
```

Field này **tự động chảy qua** `chart_data.json` (table step ghi lại nguyên văn) rồi tới
`generate()`/`surveyflow-pptx` — bỏ qua field này (hoặc set `"default"`) thì appendix tự dùng
format mặc định. Muốn đổi format cho 1 lần chạy mà không sửa `datatable.json` → dùng
`surveyflow-pptx ... --format general|default` (ghi đè, không lưu lại).

### Chọn logo khách hàng (chỉ áp dụng cho format "default")

Trước khi tạo appendix **lần đầu** cho 1 table dùng format `"default"` (table đó chưa có field
`appendix_logo` trong `datatable.json`), hỏi user:
> "Bạn muốn dùng logo khách hàng nào trong appendix?
> 1. Acecook (đã có sẵn trong template)
> 2. Khách hàng khác — bạn cần cung cấp/upload file ảnh logo
> 3. Không dùng logo khách hàng (chỉ hiện logo Q&Me)"

- Chọn **1** → bỏ qua field `appendix_logo` (hoặc set `"acecook"`) — logo Acecook đã có sẵn
  trong template, không cần làm gì thêm.
- Chọn **2** → hỏi user đường dẫn file ảnh logo (PNG/JPG) → ghi **đường dẫn đó** vào
  `appendix_logo`.
- Chọn **3** → ghi `"appendix_logo": "none"`.

```json
{
  "title": "SURVEY_NAME - Data Table",
  "sub_title": "Appendix",
  "tables": [...],
  "is_appendix": true,
  "appendix_logo": "output/SURVEY_NAME/assets/customer_logo.png",
  "banner": [...],
  "stub": [...]
}
```

Field này cũng **tự động chảy qua** `chart_data.json` rồi tới `generate()`/`surveyflow-pptx` —
chỉ cần hỏi **một lần**; các lần chạy appendix sau tự dùng lại logo đã lưu, không hỏi lại, trừ
khi user yêu cầu đổi logo. Muốn đổi logo cho 1 lần chạy mà không sửa `datatable.json` → dùng
`surveyflow-pptx ... --logo acecook|none|path/to/logo.png` (ghi đè, không lưu lại). Logo
Q&Me (bên trái) **không bao giờ** bị thay đổi bởi field này — chỉ logo bên phải (khách hàng)
mới bị swap/xoá.

**Format "default" khác "general" ở đâu:**
- Dùng slide layout công ty ("use_dz") thay vì slide trắng — background/thanh màu tự có sẵn từ
  template, không cần vẽ thủ công.
- Title/footer/số trang là **placeholder thật** của template (không phải textbox tự vẽ).
- Font Segoe UI (thay Arial), bảng màu chart riêng — 5 màu thật của brand
  (`156082/0F9ED5/A6A6A6/843C0C/4EA72E`) + 15 sắc tint/shade suy ra từ 5 màu đó, tổng 20 màu
  (tránh trùng màu khi câu có >5 lựa chọn — bug thực tế đã xảy ra, VD câu 8 lựa chọn).
- Chỉ có 2 layout: donut+stack (SA Ordinal/thang đo) và 3-cột bar ngang (dùng chung cho MA
  **và** SA-Nominal nhiều lựa chọn — template chưa có style cột dọc riêng).
- Logo Q&Me (trái) + logo khách hàng (phải, mặc định Acecook) — xem "Chọn logo khách hàng" ở
  trên. `"general"` không có logo nào (blank layout).
- Ô tag góc trên-trái tự đổi theo `--lang` của lần chạy table gần nhất (`lang` được ghi kèm vào
  `chart_data.json`): `"PHỤ LỤC"` nếu `--lang vi`, `"APPENDIX"` nếu `--lang en`. Không cần cấu hình
  gì thêm — chỉ cần chạy table đúng `--lang` mong muốn trước khi tạo appendix. Muốn ghi đè text
  khác (VD số chương khác "08") → `surveyflow-pptx ... --default-section-label "08 | PHỤ LỤC"`.
- Asset nằm ở `surveyflow/steps/appendix/appendix_templates/default_template.pptx` +
  `surveyflow/steps/appendix/chart_templates_default/{bar,donut,stacked}.xml`. Muốn đổi màu/font
  → sửa file gốc (chưa strip slide) trong PowerPoint → chạy lại
  `tools/extract_chart_templates_default.py path/to/your_unstripped_source.pptx` (xem docstring
  trong file đó — bản bundle trong repo đã bị xoá hết slide mẫu, không dùng lại được để extract).
  Logo Q&Me/khách hàng là 2 picture cố định trong slide master của `default_template.pptx`, tên
  shape `QMeLogo`/`CustomerLogo` — đổi logo Q&Me mặc định (hiếm khi cần) phải sửa trực tiếp file
  này trong PowerPoint, giữ nguyên tên shape `QMeLogo`.

### Tiêu đề slide tuỳ chỉnh — field "title" trong stub

Mỗi stub entry trong `datatable.json` có thể có thêm field **`"title"`** (mặc định không có /
`null`) — tiêu đề ngắn gọn hiển thị to trên đầu slide, thay cho label gốc (thường dài, đúng
nguyên văn câu hỏi khảo sát). Field footer/Q-label ở cuối slide **luôn** vẫn hiện label gốc đầy
đủ (không đổi) — `title` chỉ thay tiêu đề lớn.

**Thứ tự fallback** khi table step tạo `chart_data.json`: `"title"` của stub (nếu set) →
`title_i18n[lang]` của câu đó trong `metadata.json` (đã điền tự động ở Step 3a, không cần làm gì
thêm) → label gốc rút gọn (nếu cả 2 đều `null`). Vì vậy trong đa số trường hợp **không cần** ghi
`"title"` thủ công — chỉ cần khi muốn 1 tiêu đề KHÁC với `title_i18n` đã có (VD Step 3a tóm tắt
chưa ưng ý, hoặc muốn nhấn mạnh khía cạnh khác của câu hỏi).

```json
{ "question": "F7", "label": "What could make you CONSIDER TRYING economy instant noodles? (Multiple)",
  "title": "Reasons to try economy noodles", "stats": ["base", "percent"] }
```

> ⚠️ **KHÔNG tự thêm tiền tố `"{sub_title} - "` (tên table/brand) vào `title`** — dù table đó
> nằm trong 1 deck gộp nhiều brand (`is_appendix` nhiều table). Group/brand name đã hiển thị
> sẵn qua PowerPoint Section (xem `_add_pptx_sections`) mỗi khi deck có ≥2 group — lặp lại
> trong từng `title` (VD `"Hảo Hảo - Giới tính"` thay vì `"Giới tính"`) là dư thừa, cùng lý do
> divider slide theo group đã bị bỏ (xem `_add_pptx_sections` docstring). Bug thực tế đã xảy ra
> ở `VN8971 - Acecook DBA` — cả 11 table đều bị ghi `title` kiểu `"{sub_title} - ..."`, đã dọn
> lại bằng cách xoá tiền tố này khỏi mọi stub entry.

**Khi Claude thêm 1 câu vào stub** (Step 4 hoặc Workflow B) và label câu đó **dài/khó đọc làm
tiêu đề slide** (câu hỏi đầy đủ, nhiều mệnh đề phụ, ghi chú kỹ thuật kiểu "(Multiple)", "SHOW
ANSWER OF..."), Claude tự đề xuất 1 bản tóm tắt ngắn gọn (5-10 từ, giữ đúng ý câu hỏi, bỏ phần
kỹ thuật/lặp) và hỏi user xác nhận trước khi ghi vào `title`:
> "Label câu F7 khá dài để làm tiêu đề slide. Đề xuất title: 'Reasons to try economy noodles'.
> Bạn dùng đề xuất này, tự nhập title khác, hay giữ nguyên label làm tiêu đề?"

Không tự động ghi `title` mà không hỏi — khác với các auto-rule khác (mean/factor/T2B...), tóm
tắt ngôn ngữ tự nhiên là chủ quan nên luôn cần user xác nhận. Label ngắn/rõ ràng sẵn (đã ≤ ~8 từ)
thì không cần đề xuất `title`, để `null` là đủ.

**Câu matrix/ranking** — 1 field `title` trên stub entry cha (VD `"question": "Q17"`) áp dụng cho
**toàn bộ** slide row/rank sinh ra từ câu đó (VD `Q17_R1`, `Q17_R2`, ..., `Q27_2-Rank 1`); không
set title riêng theo từng row/rank được. **row_group** (`items: [...]`) → set `title` trên từng
item riêng nếu cần khác nhau theo từng câu con.

**Confirm trước khi chạy** (giống chạy pipeline) — sau khi đã hỏi + lưu `is_appendix` nếu chưa
table nào được đánh dấu (xem "Hỏi user muốn chạy bảng nào"):
> "Tôi sẽ tạo appendix PPTX (bảng: {All / tên các table đã đánh dấu is_appendix}, format
> {default/general}, logo {Acecook/khách hàng khác/không dùng}) từ chart_data.json. Bạn xác nhận
> không?"

**Cách 1 — chạy kèm pipeline** (table + appendix trong 1 lệnh):
```bash
python run_pipeline.py \
  --output-dir output/SURVEY_NAME \
  --version    v1 \
  --appendix
```

**Cách 2 — chạy riêng** (sau khi đã có chart_data.json):
```bash
surveyflow-pptx \
  output/SURVEY_NAME/v1/chart_data.json \
  output/SURVEY_NAME/v1/slides.pptx
```

Tuỳ chọn `surveyflow-pptx`: `--list-groups` (in JSON **mọi table** kèm `table_index`/`sub_title`/
`is_appendix` hiện tại rồi thoát, không tạo pptx — dùng để hỏi user "chạy bảng nào", xem "Hỏi user
muốn chạy bảng nào"), `--tables idx1,idx2` (override 1 lần: chạy đúng các `table_index` đã
`is_appendix: true` được chọn, không sửa file, vẫn group như bình thường nếu ≥2), `--table N` (ghi
đè bằng đúng 1 table_index cụ thể, bỏ qua toàn bộ cơ chế `is_appendix`/group), `--start-page N`
(số trang bắt đầu), `--format general|default` (ghi đè `appendix_format` của table cho riêng lần
chạy này, không lưu lại vào `datatable.json`), `--logo acecook|none|path/to/logo.png` (ghi đè
`appendix_logo` của table cho riêng lần chạy này, chỉ áp dụng với `--format default`, không lưu
lại vào `datatable.json`).

**Chart type tự suy ra** từ `chart_data.json`:
- `donut_stacked` (SA ≤5 / Likert) → donut Total + 100%-stacked breakdown bên phải
- `bar_horizontal` (MA) → bar ngang Total + các cột breakdown bên phải
- `bar_vertical` (SA >5) → cột dọc Total + các cột breakdown bên phải
- **`NUM`** → luôn `donut_stacked` như SA-Ordinal — mỗi bin `num_quantile`/range group là 1 slice,
  giữ nguyên thứ tự bin (không sort theo %), donut hole vẫn hiện "Mean: X" (mean thật của cả câu).
- **`multiplenumber`** → luôn `donut_stacked` như SA-Ordinal, nhưng % mỗi slice tính từ **mean per
  category đã normalize** (không phải % respondents trả lời category đó — số đó không cộng đủ 100%).
  Không hiện "Mean: X" ở donut hole (không có 1 mean tổng duy nhất, mỗi category có mean riêng).

> ⚠️ **Self-contained — KHÔNG cần `documents/temp.pptx` lúc chạy.**
> Style chart nằm sẵn trong `surveyflow/chart_templates/{bar,col,donut,stacked}.xml`.
> Muốn đổi màu / font chart: sửa `documents/temp.pptx` trong PowerPoint →
> `python tools/extract_chart_templates.py documents/temp.pptx` → chạy lại generate.

> ⚠️ **NEVER recreate or rewrite `tools/generate_pptx.py` hoặc
> `tools/extract_chart_templates.py`** — luôn có sẵn trong project.
> CLI command sau `pip install surveyflow`: `surveyflow-pptx` (thay thế `python tools/generate_pptx.py`).

---

## Output structure
```
output/SURVEY_NAME/
├── mcp/                      ← MCP raw files (fetched from QMe)
│   ├── definition.json
│   └── data_export.csv       ← assembled from read_survey_data_file chunks
├── data/                     ← rawdata.csv + metadata.json (generated from mcp/, reused)
│   ├── rawdata.csv
│   ├── metadata.json
│   ├── ft_codelist_Q5.json   ← codelist per FT question (Claude hoặc user tạo)
│   └── ft_coded.csv          ← binary coded output cho tất cả FT questions
├── datatable/                ← Claude manages this file
│   └── datatable.json
├── v1/                       ← datatable.xlsx + chart_data.json + slides.pptx
│   ├── datatable.xlsx
│   ├── chart_data.json       ← auto-generated by table step (feeds the appendix)
│   └── slides.pptx           ← PPTX chart appendix (Workflow D)
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
| "chạy bảng matrix, brands ngang, choices dọc" | Thêm `"matrix_orientation": "horizontal"` ở cấp table item; stub chỉ gồm matrix questions cùng rows |
| "tách bảng câu thường và bảng matrix horizontal" | Tạo 2 table items: 1 stub SA/MA thường, 1 stub matrix với `matrix_orientation: "horizontal"` |
| "chỉ hiện 4 brands, nhóm còn lại thành Others" | Thêm `matrix_rows` với `row_code` cho từng brand riêng, `row_codes` cho nhóm Others |
| "ẩn brand X khỏi bảng horizontal" | Xoá row_code đó ra khỏi `matrix_rows` (hoặc không liệt kê nó) |
| "compare demographics within each brand" | Thêm `"sig_direction": "columns"` vào table item |
| "refresh data / lấy data mới" | Re-fetch (prepare→poll→read) → overwrite `mcp/data_export.csv` → re-run với `--export-csv` + `--force-ingestion` |
| Đính kèm file `.xlsx` có 2 sheet "Question"+"Data" | Dùng `--xlsx-input` (xem Step 2b) — surveyflow tự convert sang definition.json + data_export.csv, không cần tự viết converter |
| "code câu FT / open-ended" | Workflow C: identify FT → tạo codelist → classify → ghi `ft_coded.csv` |
| "tạo codelist cho Q5" | Workflow C Step 2: sample responses Q5 → đề xuất codelist → user confirm → save `ft_codelist_Q5.json` |
| "user cung cấp codelist" | Workflow C Step 2: dùng codelist của user, không tự generate |
| "thêm FT coded vào datatable" | Workflow B: thêm `{Q_label}_c{code}` columns vào stub (treat như MA question) |
| "tạo appendix PPTX / chạy slides" | Workflow D: `--list-groups` để xem MỌI table + trạng thái `is_appendix` hiện tại; nếu chưa table nào đánh dấu → **hỏi user chạy bảng nào + option "All"** (xem "Hỏi user muốn chạy bảng nào") rồi ghi `is_appendix` trước khi chạy; format tự dùng `"default"` (không cần hỏi); nếu table chưa có `appendix_logo` thì hỏi user chọn logo Acecook/khách khác/không dùng trước (xem "Chọn logo khách hàng") |
| "tạo appendix riêng theo từng brand/nhãn hàng" | Tạo nhiều table item trong `datatable.json`, mỗi item có `"is_appendix": true` + `sub_title` tự do (VD "Hảo Hảo", "Omachi") — appendix tự gộp thành 1 file, mỗi brand 1 group + 1 PowerPoint Section. Nếu cần lọc respondent theo brand cho cả table → dùng field **`"filter"`** cấp table (xem "Filter cấp table") thay vì lọc từng banner column |
| "đổi bảng nào chạy appendix" | Hỏi lại (xem "Hỏi user muốn chạy bảng nào") → sửa `is_appendix` của các table liên quan (thêm/xoá `true`) trong `datatable.json` |
| "appendix 1 brand không ra slide nào dù xlsx đủ sheet" | Kiểm tra `chart_data.json` của table đó: nếu `"total": {}` cho mọi câu → table đang lọc Total theo cách cũ (`groups`/`conditions`) mà thiếu `"is_total": true`, hoặc `"filter"` cấp table bị sai `question`/`codes` — xem "Filter cấp table" và "Total có điều kiện" |
| "chỉ chạy appendix cho 1-2 brand cụ thể, không phải hết (không đổi lưu)" | `surveyflow-pptx ... --tables idx1,idx2` với `table_index` lấy từ `--list-groups` — override 1 lần, không sửa `datatable.json` |
| "chỉ xuất appendix cho 1 table cụ thể, không group" | `surveyflow-pptx ... --table N` (N = `table_index`) — bỏ qua cơ chế `is_appendix`/group, render đúng 1 table |
| "appendix theo format general / bỏ style công ty" | Ghi `"appendix_format": "general"` vào table item trong `datatable.json` (xem "Format appendix — default hay general") rồi chạy appendix bình thường — lưu 1 lần, các lần sau tự dùng lại |
| "đổi appendix về mặc định / bỏ format general" | Xoá field `appendix_format` (hoặc set `"default"`) khỏi table item đó trong `datatable.json` |
| "dùng logo khách hàng khác cho appendix" | Hỏi user đường dẫn file ảnh logo → ghi đường dẫn đó vào `"appendix_logo"` của table item trong `datatable.json` |
| "appendix không cần logo khách hàng / chỉ logo Q&Me" | Ghi `"appendix_logo": "none"` vào table item đó |
| "quay lại logo Acecook mặc định" | Xoá field `appendix_logo` (hoặc set `"acecook"`) khỏi table item đó |
| "label câu này dài quá, rút gọn tiêu đề slide" | Đề xuất bản tóm tắt ngắn (5-10 từ) → user xác nhận → ghi vào `"title"` của stub entry đó |
| "bỏ title tuỳ chỉnh, dùng lại label gốc" | Xoá field `"title"` (hoặc set `null`) khỏi stub entry đó |
| "phân loại câu SA Ordinal/Nominal" | Step 3b: Claude đọc metadata.json, tự phân loại từng câu SA **và Matrix_SA**, ghi field `scale_class` |
| "câu này sao không tự thêm mean" | Kiểm tra `scale_class` của câu đó trong metadata.json — chỉ Ordinal mới tự thêm `mean` |
| "sao câu Ordinal này không có T2B/NPS" | Step 3c: kiểm tra range code của thang đo — chỉ 1-5 tự thêm T2B/B2B, chỉ 1-10/0-10 tự thêm NPS |
| "thang đo bị đảo (không phải tần suất), mean tính sai" | Step 3c: kiểm tra `scale_high_code` trong metadata.json, thêm `factor` cho từng choice theo công thức mirror |
| "câu tần suất sao không tự có mean" | Đúng — Step 3c cố ý KHÔNG tự thêm mean/factor cho câu tần suất, phải hỏi user trước (xem ngoại lệ trong Step 3c) |
| "tự nhập factor cho câu tần suất" | Ghi đúng giá trị user đưa vào `choices` (per-choice), giữ `scale_class: "Ordinal"`, áp T2B/B2B/NPS theo range như bình thường |
| "bỏ qua/ignore mean cho câu tần suất" | Sửa `scale_class` câu đó thành `"Nominal"` trong metadata.json (Matrix_SA: cả entry cha + `sub_questions`), stub giữ `["base","percent"]` |
| "thêm câu ranking vào stub / gộp slide ranking lại" | Tách thành 2 stub entries: Top 3 (`ranking_mode: "rank_dist", ranking_top_n: 3`) + Overall (`ranking_mode: "any_rank"`) — xem mục "ranking — mặc định Top 3 + Overall" |
| "thêm câu NUM vào stub" | Tự động thêm `"mean"` vào `stats` + `"num_quantile": 4` — không cần hỏi/xác nhận trước |
| "chia N nhóm thay vì 4" | Sửa giá trị `num_quantile` trong stub entry của câu NUM đó |
| "group NUM theo range cố định (VD thập kỷ tuổi)" | Bỏ `num_quantile`, thay bằng `"choices"` tĩnh (group + hidden) — xem mục NUM trong Stub rules |
| "không group tuổi/số, giữ từng giá trị" | Bỏ cả `"choices"` lẫn `num_quantile` khỏi stub entry của câu NUM — mỗi giá trị hiện riêng |

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
Required: job_id (integer)
```
Returns: `{ status ("pending"|"processing"|"ready"|"error"), retry_after_seconds, ... }`  
Poll mỗi `retry_after_seconds` cho đến khi `status == "ready"`.

### read_survey_data_file
```
Required: job_id  (integer)
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
