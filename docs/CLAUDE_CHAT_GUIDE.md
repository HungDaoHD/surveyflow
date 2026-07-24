# Hướng dẫn sử dụng SurveyFlow + FieldCheck

> Dành cho Research / Data team — không cần biết lập trình

---

## Công cụ này làm được gì?

Bạn chỉ cần nhắn tên survey. Claude sẽ tự động:

1. Tìm survey trên Fieldcheck / QMe
2. Tải toàn bộ câu trả lời về
3. Kiểm tra xem có respondent nào trả lời sai logic không *(tuỳ chọn)*
4. Hỏi bạn chọn cột và hàng cho bảng
5. Xuất file **datatable.xlsx** hoàn chỉnh

Mỗi lần thay đổi bảng, chỉ cần nhắn tin — Claude cập nhật và xuất file mới, **không xoá file cũ**.

---

## Thiết lập một lần

Trước khi dùng lần đầu, cần cài 2 thứ. Nhờ team tech hỗ trợ nếu chưa có.

---

### 1. Skill `fieldcheck-dp`

**Skill là gì?** Là bộ hướng dẫn nạp vào Claude, giúp Claude hiểu đúng quy trình xử lý survey của team — từ cách tải data, kiểm tra chất lượng, đến cách tạo bảng.

**Cách cài:**
1. Mở **Claude Code** (ứng dụng desktop)
2. Vào **Settings** → **Skills**
3. Nhấn **Add skill** → chọn file `fieldcheck-dp.skill`
4. Xong — Claude tự nhận diện khi bạn nhắc đến survey

> Nếu chưa có file `fieldcheck-dp.skill`, liên hệ team tech để lấy.

---

### 2. MCP Fieldcheck (kết nối với QMe)

**MCP là gì?** Là cầu nối cho phép Claude truy cập thẳng vào hệ thống Fieldcheck / QMe — tự tìm survey, tự tải data mà không cần bạn làm gì thêm.

**Không có MCP:** Claude không thể tải data từ QMe, bạn phải tự export file rồi upload lên.  
**Có MCP:** Chỉ cần nhắn tên survey — Claude tự làm tất cả.

**Cách kết nối:**
1. Mở **Claude Code** → **Settings** → **Connections** (hoặc **MCP Servers**)
2. Thêm MCP với thông tin do team tech cung cấp
3. Đăng nhập tài khoản Fieldcheck khi được yêu cầu

> Sau khi kết nối, bạn sẽ thấy biểu tượng Fieldcheck trong danh sách kết nối của Claude.

---

## Bắt đầu sử dụng

Sau khi cài xong, mở Claude Code và nhắn lệnh theo cú pháp:

> **`/fieldcheck-dp chạy survey VN8966`**  
> **`/fieldcheck-dp làm bảng cho VN8894 - Express`**

Phần `/fieldcheck-dp` ở đầu là bắt buộc — đây là cách gọi đúng skill. Nếu thiếu, Claude sẽ không biết dùng quy trình này.

> 💡 Sau lần đầu khởi động, các tin nhắn tiếp theo trong cùng hội thoại **không cần** gõ lại `/fieldcheck-dp` — chỉ cần nhắn bình thường như: *"thêm câu Q8"*, *"bật sig test"*, *"chạy quality check"*…

Claude hiển thị tiến trình để bạn theo dõi:

```
📋 Pipeline: VN8966
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ 1. Tìm survey       — tìm thấy: VN8966 LIPOVITAN
✅ 2. Tải data         — 450 responses
✅ 3. Xử lý data       — xong
⏳ 4. Chọn bảng       — đang hỏi...
⬜ 5. Xuất file Excel
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## Bước 1–3: Claude tự làm

Bạn không cần làm gì trong 3 bước đầu. Claude tự tìm survey, tải data và xử lý.

**Nếu survey đã chạy trước đây**, Claude hỏi:
> *"Data của VN8966 đã có sẵn. Dùng data cũ hay tải lại từ QMe?"*

- Nhắn **"dùng cũ"** hoặc **"ok"** → bỏ qua bước tải, nhanh hơn
- Nhắn **"tải lại"** → tải data mới nhất từ QMe

---

## Bước 3b: Kiểm tra chất lượng data *(tuỳ chọn)*

Sau khi xử lý data, Claude hỏi:

> *"Bạn có muốn kiểm tra chất lượng data trước khi tạo bảng không?*  
> *Tôi sẽ rà soát toàn bộ 450 người xem có ai trả lời sai logic không."*

**Nên chạy khi:**
- Survey có nhiều câu điều kiện (câu hiện/ẩn theo câu trả lời trước)
- Cần kiểm tra trước khi giao client
- Muốn biết data sạch hay chưa

Nhắn **"chạy kiểm tra"** hoặc **"bỏ qua"** để tiếp tục.

### Kết quả kiểm tra

Claude tóm tắt ngay trong chat:

```
📊 Kết quả kiểm tra — VN8966
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Tổng số người     : 450
Có vấn đề         : 12 người (2.7%)

Loại vấn đề phát hiện:
  ❌  3 người  — bỏ qua câu hỏi bắt buộc
  ⚠️   5 người  — không trả lời câu được dẫn đến
  🔍  2 người  — trả lời câu lẽ ra không được hiển thị
  💥  2 người  — câu trả lời mâu thuẫn nhau

Câu có vấn đề nhiều nhất:
  Q15 (Mức độ sử dụng) — 6 người
  Q8  (Loại sản phẩm)  — 3 người
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

**Ý nghĩa từng loại:**

| Loại vấn đề | Ý nghĩa | Nên làm gì |
|---|---|---|
| Bỏ qua câu bắt buộc | Câu luôn phải trả lời nhưng bị bỏ trống | Cần xem xét, có thể loại profile |
| Không trả lời câu được dẫn đến | Theo logic, câu đó phải hiện nhưng không có câu trả lời | Có thể lỗi survey — báo team |
| Trả lời câu không được hiển thị | Theo logic, câu không nên hiện nhưng vẫn có câu trả lời | Lỗi routing QMe — báo team |
| Câu trả lời mâu thuẫn | Hai câu trả lời không thể đúng cùng lúc | Nên loại trước khi chạy bảng |

**Hỏi thêm về kết quả:**
> *"Xem chi tiết câu Q15"*  
> *"Người nào bị nhiều lỗi nhất?"*  
> *"Chỉ xem những người có câu trả lời mâu thuẫn"*

---

## Bước 4: Chọn bảng — Claude hỏi trong chat

Claude hỏi lần lượt **3 câu** ngay trong chat. Trả lời xong mỗi câu, Claude mới hỏi câu tiếp.

### Câu 1 — Banner (cột tiêu đề ngang)

Claude liệt kê các câu phân nhóm respondents từ survey, ví dụ:
> *"Banner gồm những câu nào? (Total luôn có sẵn)*  
> *S3 (Giới tính), S5 (Độ tuổi), S7 (Thu nhập)*  
> *Nhập số câu cách nhau bằng dấu phẩy:"*

→ Trả lời ví dụ: `S3, S5` hoặc `S3, S5, S7`

### Câu 2 — Stub (câu hỏi theo hàng)

Claude liệt kê tất cả câu hỏi của survey:
> *"Stub gồm những câu nào?*  
> *- Nhập "all" để lấy tất cả*  
> *- Hoặc nhập: Q1, Q5, Q8..."*

→ Trả lời `all` hoặc liệt kê cụ thể

### Câu 3 — Tiêu đề bảng

> *"Tiêu đề bảng? (Enter để dùng mặc định: VN8966 - Data Table)"*

→ Nhập tên hoặc nhấn Enter để dùng mặc định

Sau khi trả lời 3 câu, Claude lưu cấu hình và chuẩn bị chạy.

> 💡 Claude mặc định tạo 2 sheet: **Count** và **Percentage**. Sig test chỉ thêm khi bạn yêu cầu (xem phần "Thay đổi bảng sau khi chạy").

---

## Bước 4b: Chọn ngôn ngữ output *(nếu survey đa ngôn ngữ)*

Nếu survey có nhiều ngôn ngữ, Claude hỏi thêm:
> *"Survey này có 2 ngôn ngữ: vi (tiếng Việt) và en (tiếng Anh).*  
> *Bạn muốn xuất bảng bằng ngôn ngữ nào?"*

→ Nhắn `vi` hoặc `en`

Nếu survey chỉ có 1 ngôn ngữ, bước này được bỏ qua tự động.

---

## Bước 5: Nhận file Excel

Claude báo khi xong:

```
✅ Datatable xong!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 output/VN8966/v1/datatable.xlsx
📊 Sheets: General - Count  |  General - Pct
👥 450 respondents
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Bạn muốn thay đổi gì không? (thêm banner, thêm câu, bật sig test, tạo PPTX...)
```

Mở thư mục `output/VN8966/v1/` để lấy file.

---

## Bước 6 (tuỳ chọn): Tạo PPTX Chart Appendix

Sau khi có bảng Excel, bạn có thể yêu cầu Claude tạo bộ slide biểu đồ:

> *"Tạo appendix PPTX"*  
> *"Chạy slides"*

Claude sẽ hỏi xác nhận rồi tạo file:
```
output/VN8966/v1/slides.pptx
```

Mỗi câu hỏi → 1 slide, biểu đồ **editable** trong PowerPoint, style khớp template công ty.

**Tạo appendix riêng theo từng brand/nhãn hàng:**

> *"Tạo appendix riêng cho từng brand"*  
> *"Base của bảng Hảo Hảo chỉ tính người chọn Hảo Hảo ở Q2"*

Claude gộp tất cả brand vào **1 file PPTX**, mỗi brand là 1 Section riêng trong PowerPoint (Slide
Sorter/Outline view) để dễ nhảy nhanh giữa các brand khi review — không tốn thêm slide chia trang.

Nếu có nhiều bảng mà chưa rõ chạy bảng nào, Claude hỏi:
> *"Bạn muốn chạy appendix cho bảng nào? Total, Hảo Hảo, Mì Đệ Nhất... (1, vài tên, hoặc 'All')"*

Câu trả lời được lưu lại — lần sau tự dùng lại, không hỏi lại trừ khi bạn muốn đổi.

---

## Thay đổi bảng sau khi chạy

Chỉ cần nhắn — Claude sửa và tạo file mới *(v2, v3…)*. **File cũ vẫn còn nguyên.**

### Cột tiêu đề (banner)

| Bạn nhắn | Claude làm |
|---|---|
| "Thêm câu Income vào cột tiêu đề" | Thêm Income vào banner |
| "Bỏ cột Độ tuổi" | Xoá Age khỏi banner |
| "Thêm cột Total cho từng khu vực" | Thêm tổng trước mỗi nhóm Area |
| "Dùng Brand làm cột, tất cả brands từ Q17" | Thêm Brand matrix header |

### Câu hỏi trong bảng (stub) và thống kê

| Bạn nhắn | Claude làm |
|---|---|
| "Thêm câu Q8 vào bảng" | Thêm Q8 |
| "Bỏ Q15 ra khỏi bảng" | Xoá Q15 |
| "Thêm điểm trung bình cho Q36" | Thêm Mean cho Q36 |
| "Thêm Top 2 Box codes 4 và 5 cho Q8" | Thêm T2B |
| "Thêm tất cả câu còn lại" | Thêm hết câu chưa có trong bảng |

### Loại bảng và định dạng

| Bạn nhắn | Claude làm |
|---|---|
| "Bật kiểm định thống kê (sig test)" | Thêm sheet Sig test |
| "Tắt sig test" | Tắt sheet Sig test |
| "Chỉ giữ sheet Percentage" | Bỏ sheet Count |
| "Hiển thị 1 chữ số sau dấu phẩy" | Đổi định dạng % thành 35.2% |
| "Xuất bảng tiếng Anh" | Chạy lại với nhãn tiếng Anh |
| "Xuất bảng tiếng Việt" | Chạy lại với nhãn tiếng Việt |

### Cập nhật data

| Bạn nhắn | Claude làm |
|---|---|
| "Tải data mới nhất từ QMe" | Fetch lại + chạy lại toàn bộ |
| "Bao gồm cả người chưa duyệt (pending)" | Thêm pending profiles vào data |

---

## Khi không tải được data từ QMe

Đôi khi kết nối QMe bị treo. Claude sẽ thông báo:
> *"Tải data bị gián đoạn. Bạn có thể export file từ Fieldcheck và upload vào đây."*

**Cách xử lý:**
1. Mở **Fieldcheck** → tìm survey → **Export** → tải file `.zip` về máy
2. Kéo thả file `.zip` vào cửa sổ chat
3. Claude tự xử lý và tiếp tục

---

## Câu nói nhanh hay dùng

**Lần đầu trong hội thoại** — cần có `/fieldcheck-dp` ở đầu:
```
/fieldcheck-dp chạy survey VN8966
/fieldcheck-dp làm bảng cho VN8894 - Express
/fieldcheck-dp chạy quality check VN8966
```

**Các lần tiếp theo trong cùng hội thoại** — nhắn bình thường:
```
chạy kiểm tra chất lượng data
xem chi tiết câu Q15
người nào bị nhiều lỗi nhất?
thêm [tên câu] vào cột tiêu đề
thêm [tên câu] vào bảng
bỏ [tên câu] khỏi bảng
thêm điểm trung bình cho [câu]
thêm Top 2 Box cho [câu]
bật / tắt sig test
chỉ giữ sheet Percentage
xuất tiếng Anh / xuất tiếng Việt
tải lại data mới
tạo appendix PPTX / chạy slides
tạo appendix riêng cho từng brand
file kết quả ở đâu?
```

---

## File output ở đâu?

Tất cả file nằm trong thư mục `output/` trong cùng thư mục với Claude:

```
output/
└── VN8966/
    ├── v1/
    │   ├── datatable.xlsx         ← Bảng lần đầu
    │   ├── chart_data.json        ← Tự sinh kèm datatable.xlsx
    │   └── slides.pptx            ← PPTX appendix (nếu đã tạo)
    ├── v2/
    │   ├── datatable.xlsx         ← Sau lần chỉnh đầu tiên
    │   └── ...
    └── quality/
        └── flagged_profiles.csv   ← Danh sách người có vấn đề (nếu đã kiểm tra)
```

Mỗi lần thay đổi bảng → Claude tạo thêm thư mục mới (v2, v3…). **Các file cũ không bị xoá.**

---

## Câu hỏi thường gặp

**File Excel nằm ở đâu?**  
→ Nhắn Claude: *"file ở đâu?"* — Claude sẽ chỉ đường dẫn chính xác.

**Thay đổi bảng có mất file cũ không?**  
→ Không. Mỗi lần thay đổi Claude tạo v2, v3… file v1 vẫn còn nguyên.

**Kiểm tra chất lượng data có bắt buộc không?**  
→ Không. Nhưng nên chạy nếu survey có nhiều câu hiện/ẩn theo điều kiện, hoặc trước khi giao file cho client.

**Bảng ra thiếu người / 0 người?**  
→ Nhắn: *"bao gồm cả người chưa duyệt (pending)"* — có thể data chỉ có profile pending chưa được approve.

**Tải data từ QMe bị lỗi?**  
→ Export file zip từ Fieldcheck rồi kéo vào chat — Claude sẽ nhận và tiếp tục tự động.

**Claude có tự sửa file mà không hỏi không?**  
→ Không. Claude luôn báo trước khi tải lại data hoặc ghi đè file quan trọng. Với thay đổi bảng thông thường, Claude sửa thẳng và thông báo kết quả (không hỏi từng bước nhỏ để tránh mất thời gian).

**Muốn bảng tiếng Anh?**  
→ Nhắn: *"xuất bảng tiếng Anh"* bất cứ lúc nào — Claude chạy lại với nhãn tiếng Anh.

**Appendix tách được theo từng brand/nhãn hàng không?**  
→ Có. Nhắn *"tạo appendix riêng cho từng brand"* — Claude gộp tất cả brand vào 1 file PPTX, mỗi
brand là 1 Section riêng trong PowerPoint để dễ nhảy nhanh giữa các brand khi review.

**Muốn thêm sig test?**  
→ Nhắn: *"bật sig test"* — Claude thêm sheet Sig test và chạy lại.

**Tôi muốn xem cấu hình bảng hiện tại thì làm sao?**  
→ Nhắn: *"cho tôi xem config bảng hiện tại"* — Claude tóm tắt banner, stub, và loại bảng đang dùng.
