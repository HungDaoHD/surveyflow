# Hướng dẫn sử dụng SurveyFlow + FieldCheck

> **Dành cho:** Research team / Data team  
> **Yêu cầu:** Claude Code (desktop app) đã cài skill `surveyflow` và `fieldcheck-dp`

---

## Mục lục

1. [Tổng quan](#tổng-quan)
2. [Cài đặt một lần](#cài-đặt-một-lần)
3. [Workflow A — Chạy survey lần đầu](#workflow-a--chạy-survey-lần-đầu)
4. [Workflow B — Thay đổi bảng (đã có data)](#workflow-b--thay-đổi-bảng)
5. [Field Check — Kiểm tra logic trả lời](#field-check)
6. [Upload Excel template để thiết kế bảng](#excel-template)
7. [Các yêu cầu thông thường](#các-yêu-cầu-thông-thường)
8. [Cấu trúc folder output](#cấu-trúc-folder-output)
9. [Câu hỏi thường gặp](#câu-hỏi-thường-gặp)

---

## Tổng quan

SurveyFlow là pipeline tự động xử lý data từ **QMe** thành 3 file:

| File | Nội dung |
|---|---|
| `rawdata.csv` | Data thô dạng số, 1 row = 1 respondent |
| `metadata.json` | Danh sách câu hỏi, nhãn, choice codes |
| `datatable.xlsx` | Bảng crosstab hoàn chỉnh (Count / Pct / Sig) |

**Claude** là người vận hành pipeline. Bạn chỉ cần nói bằng ngôn ngữ tự nhiên — Claude sẽ tự fetch data, thiết kế bảng, và chạy pipeline.

---

## Cài đặt một lần

### 1. Cài Python package

```bash
pip install surveyflow
```

Kiểm tra:
```bash
python -c "import surveyflow; print(surveyflow.__version__)"
```

### 2. Cài skill trong Claude Code

Mở Claude Code → Settings → Skills → Add skill file:
- `skills/surveyflow.skill` — pipeline bảng crosstab
- `skills/fieldcheck-dp.skill` — kiểm tra logic trả lời

Sau khi cài, Claude tự nhận diện khi bạn nói về survey.

### 3. Kết nối QMe MCP

Claude Code cần được kết nối với QMe MCP server để fetch data trực tiếp từ QMe.  
Liên hệ team tech để được cấp config MCP.

---

## Workflow A — Chạy survey lần đầu

### Bước 1: Tìm survey

Nói với Claude:
> *"Chạy survey [tên survey]"*  
> *"Tìm survey VN8963 LIPOVITAN"*

Claude sẽ tìm `survey_id` trong QMe.

---

### Bước 2: Fetch data từ QMe

Claude tự động fetch nếu chưa có data. Nếu data đã tồn tại, Claude sẽ hỏi:
> *"Data đã có sẵn. Dùng data cũ hay fetch lại?"*

- **Dùng data cũ** → bỏ qua fetch, dùng file đã có
- **Fetch lại** → tải data mới từ QMe (ghi đè file cũ)

> ⚠️ Claude sẽ xin xác nhận trước khi ghi đè data cũ.

---

### Bước 3: Chạy ingestion (tạo rawdata + metadata)

Claude tự chạy, không cần làm gì thêm. Kết quả:
```
output/SURVEY_NAME/data/rawdata.csv
output/SURVEY_NAME/data/metadata.json
```

---

### Bước 4: Thiết kế bảng

Claude hỏi lần lượt 3 câu:

**Câu 1 — Loại bảng:**
```
1. Count only
2. Percentage only  
3. Percentage + Sig test (90% & 95%)
4. Tất cả (Count + Pct + Sig)
```

**Câu 2 — Banner (header ngang):**  
Claude liệt kê các câu SA/MA, bạn chọn câu nào làm header.  
Ví dụ: `Q2 (Giới tính), Q3 (Độ tuổi)`

**Câu 3 — Stub (câu hỏi theo hàng):**  
```
- "all" → lấy tất cả câu
- Hoặc liệt kê: Q5, Q8, Q12, Q15...
```

> **Cách nhanh hơn:** Upload file Excel template (xem [phần 6](#excel-template))

---

### Bước 5: Chạy pipeline → xuất Excel

Claude sẽ báo:
> *"Tôi sẽ chạy pipeline v1 với config hiện tại. Bạn xác nhận chạy không?"*

Xác nhận → Claude chạy → file xuất ra:
```
output/SURVEY_NAME/v1/datatable.xlsx
```

---

## Workflow B — Thay đổi bảng

Khi đã có `datatable.xlsx` và muốn thay đổi, chỉ cần nói với Claude. Claude sẽ sửa config và chạy lại (version tăng dần: v1 → v2 → v3…).

### Ví dụ yêu cầu

| Bạn nói | Claude làm |
|---|---|
| "Thêm Income vào banner" | Thêm câu Income vào header |
| "Bỏ Q15 khỏi stub" | Xoá Q15 khỏi danh sách câu hỏi |
| "Thêm mean và std cho Q36" | Thêm thống kê mean, std vào Q36 |
| "Thêm T2B cho tất cả câu scale" | Thêm Top 2 Box cho các câu rating |
| "Chỉ giữ 1 sheet Percentage" | Tắt sheet Count và Sig |
| "Hiển thị 1 chữ số thập phân" | Thêm `decimal: 1` cho sheet Pct |
| "Thêm total cho từng banner group" | Thêm cột Total cho từng nhóm banner |
| "Nhóm Q13/Q14/Q17 theo brand" | Tạo row_group theo tên brand |
| "Brand làm header, dùng Q17" | Thêm banner_matrix từ Q17 |
| "Refresh data / fetch lại" | Tải data mới từ QMe, chạy lại toàn bộ |

---

## Field Check

Field Check kiểm tra **logic trả lời** của từng respondent — phát hiện các lỗi như:

| Loại vi phạm | Ý nghĩa |
|---|---|
| `show_condition_extra` | Trả lời câu không được hiển thị (skip logic bị bỏ qua) |
| `show_condition_missing` | Không trả lời câu đáng lẽ phải được hiển thị |
| `contradiction` | Câu trả lời mâu thuẫn với nhau |
| `always_shown_missing` | Câu luôn hiển thị nhưng bị bỏ trống |

### Cách chạy Field Check

Nói với Claude:
> *"Chạy field check cho survey [tên]"*  
> *"Kiểm tra logic trả lời survey VN8963"*  
> *"Chạy quality check"*

Claude sẽ chạy pipeline với lệnh:
```bash
python run_pipeline.py --output-dir output/SURVEY_NAME --run-quality
```

### Kết quả

Xuất ra `output/SURVEY_NAME/quality/`:

```
quality_report.json       ← tóm tắt + danh sách vi phạm
flagged_profiles.csv      ← chi tiết từng respondent bị flag
```

**`flagged_profiles.csv`** có các cột:

| Cột | Nội dung |
|---|---|
| `profile_id` | ID của respondent |
| `type` | Loại vi phạm |
| `question` | Câu hỏi liên quan |
| `detail` | Mô tả chi tiết vi phạm |
| `condition_eval` | Điều kiện được đánh giá |
| `condition_trigger` | Nhánh điều kiện thực sự trigger |

### Ví dụ kết quả

```
profile_id | type                  | question | detail
-----------|----------------------|----------|--------
task_001   | show_condition_extra  | Q15      | Answered Q15 but condition was False
task_002   | contradiction         | Q8, Q12  | Q8=1 contradicts Q12=3
task_003   | always_shown_missing  | S1       | Mandatory question S1 has no answer
```

### Hỏi Claude về kết quả

Sau khi chạy xong, bạn có thể hỏi Claude:
> *"Tóm tắt kết quả field check"*  
> *"Bao nhiêu profile bị flag? Loại nào nhiều nhất?"*  
> *"Show tôi danh sách profile bị lỗi contradiction"*

---

## Excel Template

Thay vì trả lời từng câu hỏi, bạn có thể **điền vào file Excel** rồi upload cho Claude.

### Download template

File: `tools/datatable_template.xlsx`

Template có 4 sheets:

---

### Sheet 1: Config

Cấu hình chung cho bảng.

| Key | Giá trị |
|---|---|
| title | Tên bảng (VD: "VN8963 - Data Table") |
| sub_title | Tên tab (VD: "General") |
| Count | TRUE / FALSE |
| Pct | TRUE / FALSE |
| decimal | Số chữ số thập phân (0, 1, hoặc 2) |

---

### Sheet 2: Banner

Thiết kế header ngang của bảng.

| Cột | Hướng dẫn |
|---|---|
| `question` | Tên câu (VD: Q2). Ô trống = cùng câu với dòng trên |
| `group_label` | Tên nhóm (VD: Male, Female) |
| `value` | Code(s) — 1 số hoặc nhiều số cách nhau bằng `;` (VD: `1;2;3`) |

**Ví dụ:**

| question | group_label | value |
|---|---|---|
| Q2 | Male | 1 |
| | Female | 2 |
| Q3 | 18-24 | 1;2 |
| | 25-34 | 3;4 |
| | 35+ | 5;6;7 |

> Cột Total luôn được thêm tự động — không cần điền.

---

### Sheet 3: Stub

Danh sách câu hỏi theo hàng và các thống kê cần tính.

| Cột | Hướng dẫn |
|---|---|
| `question` | Tên câu (VD: Q5, Q8, Q12) |
| `mean` | Nhập `1` nếu muốn tính mean |
| `T2B` | Nhập các code cần group, cách nhau bằng `;` (VD: `4;5`) |
| `B2B` | Tương tự T2B nhưng là Bottom 2 Box |

> ⚠️ **Nếu để trống toàn bộ sheet Stub** → Claude tự động lấy tất cả câu hỏi.

**Ví dụ:**

| question | mean | T2B | B2B |
|---|---|---|---|
| Q5 | | 4;5 | 1;2 |
| Q8 | 1 | | |
| Q12 | | 4;5 | |

---

### Sheet 4: Custom (nâng cao)

Định nghĩa filter groups dùng chung cho nhiều bảng (ví dụ: nhóm Users theo brand sử dụng).

| Cột | Hướng dẫn |
|---|---|
| `group_name` | Tên nhóm (VD: "Users") |
| `sub_label` | Tên sub-group (VD: "Brand A users") |
| `q1` | Câu hỏi điều kiện 1 |
| `codes1` | Codes cho điều kiện 1 (cách nhau bằng `;`) |
| `q2` | Câu hỏi điều kiện 2 (tùy chọn) |
| `codes2` | Codes cho điều kiện 2 |
| `logic` | `and` / `or` (mặc định: `and`) |

Để dùng Custom group trong Banner sheet, nhập `custom:GroupName` vào cột `question`.

---

### Cách upload

1. Mở `tools/datatable_template.xlsx`, điền vào
2. Lưu file
3. Trong Claude Code, gửi file kèm tin nhắn:  
   > *"Đây là config bảng cho survey VN8963. Tạo datatable.json và chạy pipeline"*

Claude sẽ:
- Parse file Excel
- Tóm tắt những gì đọc được (banner, stub, loại bảng)
- Xin xác nhận trước khi tạo và chạy

---

## Các yêu cầu thông thường

### Bảng thông thường

| Bạn nói | Claude làm |
|---|---|
| "Chạy survey [tên]" | Fetch data → ingest → hỏi thiết kế bảng → chạy |
| "Chạy lại bảng với config mới" | Sửa datatable.json + chạy version mới |
| "Xem bảng hiện tại đang config gì" | Đọc và tóm tắt datatable.json |
| "Bật sig test" | Thêm sig test cho sheet Sig |
| "Tắt sig test" | Tắt sig test |
| "Thêm sheet Count" | Thêm sheet Count vào tables |

### Banner

| Bạn nói | Claude làm |
|---|---|
| "Thêm [câu X] vào banner" | Thêm câu X với các choice codes |
| "Xoá [câu X] khỏi banner" | Xoá câu X khỏi banner |
| "Đổi tên group Male thành Nam" | Đổi label |
| "Thêm total cho từng area" | Thêm `show_total: true` cho banner Area |
| "Brand làm header, dùng Q17" | Thêm `banner_matrix` từ Q17 |
| "Nhóm International = brands 1,2,3,4" | Thêm group với `row_codes` |

### Stub

| Bạn nói | Claude làm |
|---|---|
| "Thêm [câu X] vào stub" | Thêm câu X |
| "Bỏ [câu X]" | Xoá câu X |
| "Thêm mean cho Q36" | Thêm `"mean"` vào stats của Q36 |
| "Thêm T2B 4-5 cho Q5" | Thêm `"t2b"` + choices |
| "Thêm tất cả câu" | Thêm hết câu SA/MA/Matrix |
| "Nhóm Q13/Q14/Q17 theo brand" | Tạo `row_group` |

### Field Check

| Bạn nói | Claude làm |
|---|---|
| "Chạy field check" | Chạy quality check |
| "Tóm tắt kết quả quality check" | Đọc report và tóm tắt |
| "Có bao nhiêu profile bị flag?" | Đếm và báo số |
| "Show profile bị contradiction" | Lọc và liệt kê |

---

## Cấu trúc folder output

```
output/
└── SURVEY_NAME/
    ├── mcp/                    ← Data thô từ QMe (fetch 1 lần)
    │   ├── definition.json
    │   └── data_export.csv
    ├── data/                   ← Kết quả ingestion (tái sử dụng)
    │   ├── rawdata.csv
    │   └── metadata.json
    ├── datatable/              ← Claude quản lý file này
    │   └── datatable.json
    ├── quality/                ← Kết quả field check
    │   ├── quality_report.json
    │   └── flagged_profiles.csv
    ├── v1/
    │   └── datatable.xlsx      ← Bảng version 1
    ├── v2/
    │   └── datatable.xlsx      ← Sau khi thay đổi
    └── v3/
        └── datatable.xlsx
```

---

## Câu hỏi thường gặp

**Q: Mỗi lần chạy có cần fetch data lại không?**  
A: Không. Data trong `mcp/` được tái sử dụng. Chỉ fetch lại khi QMe có response mới hoặc cần cập nhật.

**Q: Thay đổi bảng có mất data gốc không?**  
A: Không. Data trong `data/` không bao giờ bị ghi đè khi chỉ thay đổi bảng. Chỉ `datatable.json` và `v{n}/datatable.xlsx` thay đổi.

**Q: Version v1, v2, v3 khác gì nhau?**  
A: Mỗi lần user yêu cầu thay đổi bảng, Claude tăng version để giữ lịch sử. File cũ vẫn còn đó.

**Q: Sig test là gì?**  
A: Statistical significance test — kiểm định xem sự khác biệt giữa các nhóm có ý nghĩa thống kê không. SurveyFlow mặc định dùng Welch's t-test ở mức tin cậy 90% và 95%.

**Q: Field check kiểm tra gì?**  
A: Kiểm tra từng respondent có trả lời đúng theo logic hiển thị (show conditions) của survey không. Ví dụ: respondent không dùng brand X mà vẫn trả lời câu về mức độ hài lòng với brand X.

**Q: Claude có tự ý sửa file không?**  
A: Không. Claude luôn tóm tắt và xin xác nhận trước khi:
- Sửa `datatable.json`
- Chạy pipeline
- Fetch lại data từ QMe

**Q: Tôi muốn xem config bảng hiện tại thì làm sao?**  
A: Nói với Claude: *"Cho tôi xem config bảng hiện tại"* hoặc *"datatable.json đang có gì?"*

---

*Tài liệu này được quản lý cùng với codebase surveyflow. Phiên bản cập nhật nhất luôn ở `docs/USER_GUIDE.md`.*
