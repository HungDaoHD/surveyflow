# Hướng dẫn sử dụng — SurveyFlow / Fieldcheck-DP

> Công cụ xử lý dữ liệu survey từ QMe → file Excel datatable.

---

## Bắt đầu như thế nào?

Chỉ cần nói với Claude tên survey:

```
chạy survey VN8966
làm bảng cho VN8894 - Express
fetch data VN8966
```

Claude sẽ tự động thực hiện toàn bộ quy trình.

---

## Quy trình chuẩn (lần đầu)

```
1. Tìm survey trên QMe
2. Tải data về
3. Xử lý rawdata + metadata
4. Bạn chọn banner & stub trong giao diện builder
5. Chạy ra file Excel
```

Toàn bộ output lưu tại: `output/TÊN_SURVEY/`

---

## File bạn nhận được

| File | Nội dung |
|---|---|
| `v1/datatable.xlsx` | Bảng cross-tab hoàn chỉnh |
| `data/rawdata.csv` | Dữ liệu thô đã chuẩn hoá |
| `data/metadata.json` | Cấu trúc câu hỏi (dùng nội bộ) |

---

## Chọn banner & stub

Sau khi tải data xong, Claude sẽ mở **Datatable Builder** — giao diện chọn cột/hàng:

- **Banner (cột)** = câu hỏi phân nhóm respondents, ví dụ: Giới tính, Độ tuổi, Khu vực
- **Stub (hàng)** = câu hỏi muốn xem kết quả, ví dụ: Q1, Q2, Q3...

**Cách dùng builder:**
1. Kéo thả file `metadata.json` vào ô upload (Claude sẽ chỉ đường dẫn)
2. Chọn loại output: Percentage + Count / Percentage only / Count only
3. Chọn câu làm banner
4. Chọn câu làm stub
5. Nhấn **"Tạo datatable.json"** → sao chép kết quả → dán vào chat

> 💡 Muốn thêm **Sig test** sau khi có bảng? Nói với Claude: *"bật sig test"* — Claude sẽ thêm vào và chạy lại.

---

## Thay đổi bảng sau khi chạy

Chỉ cần nói tự nhiên:

| Bạn nói | Claude làm |
|---|---|
| "thêm câu Q5 vào stub" | Thêm Q5, chạy lại ra v2 |
| "bỏ Age ra khỏi banner" | Xoá banner Age, chạy lại |
| "thêm mean cho Q36" | Thêm giá trị mean vào Q36 |
| "xuất tiếng Anh" | Chạy lại với nhãn tiếng Anh |
| "thêm sig test" | Thêm sheet Sig test |
| "chỉ lấy 1 sheet percentage" | Bỏ sheet Count, giữ lại Pct |

Mỗi lần thay đổi Claude tạo version mới (v2, v3...) — file cũ vẫn giữ nguyên.

---

## Quality check — kiểm tra chất lượng data

**Đây là tính năng tuỳ chọn.** Claude sẽ hỏi bạn sau khi ingestion xong.

Khi nào nên chạy:
- Survey có nhiều câu điều kiện (show/hide logic)
- Cần kiểm tra trước khi giao client
- Muốn biết profile nào bị lỗi routing

Quality check phát hiện 4 loại vấn đề:

| Loại | Nghĩa là |
|---|---|
| `missing` | Câu hỏi luôn hiển thị nhưng respondent không trả lời |
| `routed_missing` | Câu được route đến nhưng không có câu trả lời |
| `extraneous` | Respondent trả lời câu lẽ ra không được hiển thị |
| `contradiction` | Câu trả lời vi phạm logic đặt trong survey |

Claude sẽ tóm tắt kết quả và gợi ý xử lý. Bạn cũng có thể hỏi thêm:

```
xem chi tiết câu Q5
profile nào bị lỗi nhiều nhất
chỉ xem contradiction
```

---

## Lấy data mới

Nếu QMe có thêm responses mới:

```
refresh data VN8966
lấy data mới VN8966
```

Claude sẽ fetch lại và hỏi bạn xác nhận trước khi ghi đè file cũ.

---

## Thêm profile pending

Mặc định chỉ lấy profile đã `approved`. Nếu muốn bao gồm cả `pending`:

```
bao gồm cả pending
chạy với approved và pending
```

---

## Một số câu hỏi thường gặp

**Q: File Excel output ở đâu?**
A: `output/TÊN_SURVEY/v1/datatable.xlsx` (v1 tăng lên v2, v3... mỗi lần thay đổi)

**Q: Chạy lại từ đầu có mất data cũ không?**
A: Không. File Excel cũ vẫn còn trong folder `v1/`, `v2/`... Claude chỉ tạo folder mới.

**Q: Có thể chạy nhiều survey cùng lúc không?**
A: Được — mỗi survey có folder riêng trong `output/`. Chạy lần lượt với Claude.

**Q: Builder không load được metadata.json?**
A: Đường dẫn đúng là `output/TÊN_SURVEY/data/metadata.json` — hỏi Claude để lấy đường dẫn chính xác.

**Q: Muốn bảng tiếng Anh?**
A: Nói với Claude: *"xuất tiếng Anh"* hoặc *"--lang en"* khi chạy.
