# Hướng dẫn sử dụng SurveyFlow + FieldCheck

> **Dành cho:** Research team / Data team  
> **Yêu cầu:** Claude Code (desktop app) đã cài skill `fieldcheck-dp`

---

## Mục lục

1. [Tổng quan](#tổng-quan)
2. [Cài đặt một lần](#cài-đặt-một-lần)
3. [Workflow A — Chạy survey lần đầu](#workflow-a--chạy-survey-lần-đầu)
4. [Workflow B — Thay đổi bảng (đã có data)](#workflow-b--thay-đổi-bảng)
5. [PPTX Chart Appendix](#pptx-chart-appendix)
6. [Field Check — Kiểm tra logic trả lời](#field-check)
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
- `.skill_work/fieldcheck-dp.skill` — workflow chính: fetch / ingest / check / datatable / PPTX

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

Hoặc dùng lệnh skill:
> `/fieldcheck-dp chạy survey VN8963`

Claude sẽ tìm `survey_id` trong QMe.

---

### Bước 2: Lấy data từ QMe

Claude tự động fetch nếu chưa có data. Nếu data đã tồn tại, Claude sẽ hỏi:
> *"Data đã có sẵn. Dùng data cũ hay lấy lại từ QMe?"*

- **Dùng data cũ** → bỏ qua fetch, dùng file đã có
- **Fetch lại** → tải data mới từ QMe (ghi đè file cũ)

Nếu QMe bị treo hoặc không kết nối được, Claude hỏi:
> *"Bạn muốn fetch từ QMe hay upload file zip từ Fieldcheck?"*

→ Export file `.zip` từ Fieldcheck và kéo thả vào cửa sổ chat.

> ⚠️ Claude sẽ xin xác nhận trước khi ghi đè data cũ.

---

### Bước 3: Chạy ingestion (tạo rawdata + metadata)

Claude tự chạy, không cần làm gì thêm. Kết quả:
```
output/SURVEY_NAME/data/rawdata.csv
output/SURVEY_NAME/data/metadata.json
```

Sau ingestion, Claude hỏi:
> *"Bạn có muốn chạy quality check trước khi tạo bảng không?"*

---

### Bước 3b: Quality check *(tuỳ chọn)*

Nói **"chạy quality"** để kiểm tra logic trả lời, hoặc **"bỏ qua"** để tiếp tục.

Xem chi tiết ở phần [Field Check](#field-check).

---

### Bước 4: Thiết kế bảng

Claude hỏi lần lượt **3 câu** trong chat (chờ trả lời xong mỗi câu rồi mới hỏi câu tiếp):

**Câu 1 — Banner (header ngang):**
Claude liệt kê các câu SA/MA, bạn chọn câu nào làm header.  
Ví dụ trả lời: `S3, S5` hoặc `S3, S5, S7`

**Câu 2 — Stub (câu hỏi theo hàng):**
```
- "all" → lấy tất cả câu SA/MA/Matrix
- Hoặc liệt kê: Q1, Q5, Q8, Q12...
```

**Câu 3 — Tiêu đề bảng:**  
Nhập tên hoặc nhấn Enter để dùng mặc định.

Claude tạo `datatable.json` với **default tables: Count + Pct**.  
Sig test chỉ thêm khi bạn yêu cầu ("bật sig test").

---

### Bước 4b: Chọn ngôn ngữ output

Nếu survey có **nhiều ngôn ngữ**, Claude hỏi thêm:
> *"Survey có vi (tiếng Việt) và en (tiếng Anh). Bạn muốn xuất ngôn ngữ nào?"*

Nếu chỉ có 1 ngôn ngữ, bước này được bỏ qua tự động.

---

### Bước 5: Chạy pipeline → xuất Excel

Claude sẽ báo:
> *"Tôi sẽ chạy pipeline v1. Bạn xác nhận chạy không?"*

Xác nhận → Claude chạy → file xuất ra:
```
output/SURVEY_NAME/v1/datatable.xlsx
output/SURVEY_NAME/v1/chart_data.json   ← tự động sinh, dùng cho PPTX appendix
```

---

### Bước 6 (tuỳ chọn): Tạo PPTX Chart Appendix

Sau khi có `datatable.xlsx`, bạn có thể yêu cầu:

> *"Tạo appendix PPTX"*  
> *"Chạy slides"*

Xem chi tiết ở phần [PPTX Chart Appendix](#pptx-chart-appendix).

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
| "Bật sig test" | Thêm sheet Sig test (90% & 95%) |
| "Tắt sig test" | Bỏ sheet Sig test |
| "Chỉ giữ 1 sheet Percentage" | Tắt sheet Count và Sig |
| "Hiển thị 1 chữ số thập phân" | Thêm `decimal: 1` cho sheet Pct |
| "Thêm total cho từng banner group" | Thêm cột Total cho từng nhóm banner |
| "Nhóm Q13/Q14/Q17 theo brand" | Tạo row_group theo tên brand |
| "Brand làm header, dùng Q17" | Thêm banner_matrix từ Q17 |
| "Refresh data / fetch lại" | Tải data mới từ QMe, chạy lại toàn bộ |
| "Xuất tiếng Anh" | Re-run với `--lang en` |

---

## PPTX Chart Appendix

Sau khi chạy bảng, `chart_data.json` được tự động sinh ra cùng `datatable.xlsx`. Từ đó có thể tạo bộ slide biểu đồ — mỗi câu hỏi 1 slide, **editable** trong PowerPoint.

**Cách tạo:**

> *"Tạo appendix PPTX cho survey VN8963 v1"*  
> *"Chạy slides"*

Claude hỏi xác nhận rồi chạy:
```
output/SURVEY_NAME/v1/slides.pptx
```

**Chart type tự suy ra** từ dữ liệu:
- SA với ≤5 lựa chọn / Likert → donut chart
- MA → bar ngang
- SA với >5 lựa chọn → cột dọc

---

## Field Check <a name="field-check"></a>

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
> *"Kiểm tra logic trả lời"*  
> *"Chạy quality check"*

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

### Hỏi Claude về kết quả

Sau khi chạy xong, bạn có thể hỏi Claude:
> *"Tóm tắt kết quả field check"*  
> *"Bao nhiêu profile bị flag? Loại nào nhiều nhất?"*  
> *"Show tôi danh sách profile bị lỗi contradiction"*

---

## Các yêu cầu thông thường

### Bảng thông thường

| Bạn nói | Claude làm |
|---|---|
| "Chạy survey [tên]" | Fetch data → ingest → hỏi thiết kế bảng → chạy |
| "Chạy lại bảng với config mới" | Sửa datatable.json + chạy version mới |
| "Xem bảng hiện tại đang config gì" | Đọc và tóm tắt datatable.json |
| "Bật sig test" | Thêm sheet Sig test (90% & 95%) |
| "Tắt sig test" | Tắt sheet Sig test |
| "Xuất tiếng Anh" | Chạy lại với `--lang en` |

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
| "Thêm T2B 4-5 cho Q5" | Thêm `"t2b"` vào stats của Q5 |
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
    │   ├── datatable.xlsx      ← Bảng version 1
    │   ├── chart_data.json     ← Tự sinh kèm datatable.xlsx
    │   └── slides.pptx         ← PPTX appendix (nếu đã tạo)
    ├── v2/
    │   ├── datatable.xlsx      ← Sau khi thay đổi
    │   ├── chart_data.json
    │   └── slides.pptx
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
A: Statistical significance test — kiểm định xem sự khác biệt giữa các nhóm có ý nghĩa thống kê không. SurveyFlow dùng Welch's t-test ở mức tin cậy 90% và 95%.

**Q: Sig test không xuất hiện mặc định?**  
A: Đúng. Mặc định chỉ có sheet Count và Pct. Nói *"bật sig test"* để thêm sheet Sig.

**Q: Field check kiểm tra gì?**  
A: Kiểm tra từng respondent có trả lời đúng theo logic hiển thị (show conditions) của survey không. Ví dụ: respondent không dùng brand X mà vẫn trả lời câu về mức độ hài lòng với brand X.

**Q: Claude có tự ý sửa file không?**  
A: Không. Claude luôn tóm tắt và xin xác nhận trước khi:
- Sửa `datatable.json`
- Chạy pipeline
- Fetch lại data từ QMe

**Q: Tôi muốn xem config bảng hiện tại thì làm sao?**  
A: Nói với Claude: *"Cho tôi xem config bảng hiện tại"* hoặc *"datatable.json đang có gì?"*

**Q: Bảng ra 0 rows?**  
A: Nói *"bao gồm cả người chưa duyệt (pending)"* — có thể data chỉ có profile pending chưa approved.

---

*Tài liệu này được quản lý cùng với codebase surveyflow. Phiên bản cập nhật nhất luôn ở `docs/USER_GUIDE.md`.*
