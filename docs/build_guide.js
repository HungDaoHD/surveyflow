const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, LevelFormat,
} = require("docx");
const fs = require("fs");

// ── Colours ───────────────────────────────────────────────────────────────────
const BLUE       = "1F5C99";
const BLUE_MID   = "2E75B6";
const BLUE_LIGHT = "D6E4F3";
const GRAY_BG    = "F2F2F2";
const GRAY_BD    = "CCCCCC";
const GREEN_BG   = "E8F5E9";
const YELLOW_BG  = "FFF9E6";
const WHITE      = "FFFFFF";

const DXA   = (in_) => Math.round(in_ * 1440);
const PAGE_W    = DXA(11.69);
const MARGIN    = DXA(1.0);
const CONTENT_W = PAGE_W - MARGIN * 2;   // 9026

// ── Primitive helpers ─────────────────────────────────────────────────────────
const r = (text, opts = {}) =>
  new TextRun({ text, font: "Arial", size: 22, ...opts });
const rb = (text, opts = {}) => r(text, { bold: true, ...opts });
const ri = (text, opts = {}) => r(text, { italics: true, ...opts });
const mono = (text) =>
  new TextRun({ text, font: "Courier New", size: 20, color: "333333" });

const b1 = (color = GRAY_BD) => ({ style: BorderStyle.SINGLE, size: 4, color });
const borders = (color = GRAY_BD) => ({
  top: b1(color), bottom: b1(color), left: b1(color), right: b1(color),
});
const noBorders = () => {
  const n = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: n, bottom: n, left: n, right: n };
};

// ── Block helpers ─────────────────────────────────────────────────────────────
const sp = (pt = 160) => new Paragraph({ children: [], spacing: { after: pt } });

function para(runs, opts = {}) {
  const c = Array.isArray(runs) ? runs : [r(runs)];
  return new Paragraph({ children: c, spacing: { after: 120 }, ...opts });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 30, bold: true, color: WHITE })],
    shading: { fill: BLUE_MID, type: ShadingType.CLEAR },
    spacing: { before: 360, after: 200 },
    indent: { left: DXA(0.15), right: DXA(0.15) },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: BLUE_MID })],
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE_MID, space: 2 } },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: "333333" })],
    spacing: { before: 180, after: 80 },
  });
}

function bullet(runs) {
  const c = Array.isArray(runs) ? runs : [r(runs)];
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: c,
    spacing: { after: 80 },
  });
}

function numbered(runs) {
  const c = Array.isArray(runs) ? runs : [r(runs)];
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    children: c,
    spacing: { after: 80 },
  });
}

// Single-cell coloured callout box
function callout(lines, bgColor = BLUE_LIGHT, icon = "💡") {
  const paragraphs = lines.map((line, i) => {
    const runs = Array.isArray(line) ? line : [ri(line)];
    return new Paragraph({
      children: i === 0
        ? [new TextRun({ text: icon + "  ", font: "Segoe UI Emoji", size: 22 }), ...runs]
        : runs,
      spacing: { after: i < lines.length - 1 ? 80 : 0 },
    });
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: borders("A0BFE0"),
      shading: { fill: bgColor, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: paragraphs,
    })] })],
  });
}

// Monospace grey code block
function codeBlock(lines) {
  const paragraphs = lines.map((line, i) =>
    new Paragraph({
      children: [mono(line)],
      spacing: { after: i < lines.length - 1 ? 40 : 0 },
    })
  );
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: borders("AAAAAA"),
      shading: { fill: GRAY_BG, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 220, right: 220 },
      children: paragraphs,
    })] })],
  });
}

// Data table with blue header row
function dataTable(headers, rows, colWidths) {
  const cell = (runs, isHeader, w) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: borders(),
    shading: { fill: isHeader ? BLUE : WHITE, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 130, right: 130 },
    children: [new Paragraph({
      children: Array.isArray(runs) ? runs : [new TextRun({
        text: runs, font: "Arial", size: 20,
        bold: isHeader, color: isHeader ? WHITE : "333333",
      })],
      spacing: { after: 0 },
    })],
  });

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => cell(
      [new TextRun({ text: h, font: "Arial", size: 20, bold: true, color: WHITE })],
      true, colWidths[i]
    )),
  });

  const dataRows = rows.map((row) =>
    new TableRow({
      children: row.map((cell_content, i) => {
        const runs = Array.isArray(cell_content)
          ? cell_content
          : [new TextRun({ text: cell_content, font: "Arial", size: 20, color: "333333" })];
        return cell(runs, false, colWidths[i]);
      }),
    })
  );

  return new Table({
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// FAQ item: bold question + indented answer
function faq(question, answerRuns) {
  const runs = Array.isArray(answerRuns)
    ? answerRuns
    : [r(answerRuns)];
  return [
    new Paragraph({
      children: [rb("❓  " + question, { color: BLUE_MID })],
      spacing: { before: 180, after: 60 },
    }),
    new Paragraph({
      children: [r("→  "), ...runs],
      spacing: { after: 100 },
      indent: { left: DXA(0.25) },
    }),
  ];
}

// ── Document content ──────────────────────────────────────────────────────────
const children = [

  // ── Title ──
  new Paragraph({
    children: [new TextRun({ text: "Hướng dẫn sử dụng", font: "Arial", size: 52, bold: true, color: BLUE_MID })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 120 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "SurveyFlow + FieldCheck", font: "Arial", size: 64, bold: true, color: BLUE })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
  }),
  new Paragraph({
    children: [ri("Dành cho Research / Data team — không cần biết lập trình", { size: 24, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
  }),

  // ─── 1. Làm được gì ───────────────────────────────────────────────────────
  h1("Công cụ này làm được gì?"),
  para([r("Bạn chỉ cần nhắn tên survey. Claude sẽ tự động:")]),
  numbered("Tìm survey trên Fieldcheck / QMe"),
  numbered("Tải toàn bộ câu trả lời về"),
  numbered([r("Kiểm tra xem có respondent nào trả lời sai logic không "), ri("(tuỳ chọn)")]),
  numbered("Mở giao diện để bạn chọn cột và hàng cho bảng"),
  numbered([rb("Xuất file datatable.xlsx hoàn chỉnh")]),
  sp(80),
  para([r("Mỗi lần thay đổi bảng, chỉ cần nhắn tin — Claude cập nhật và xuất file mới, "), rb("không xoá file cũ"), r(".")]),
  sp(),

  // ─── 2. Thiết lập ─────────────────────────────────────────────────────────
  h1("Thiết lập một lần"),
  para("Trước khi dùng lần đầu, cần cài 2 thứ. Nhờ team tech hỗ trợ nếu chưa có."),
  sp(80),

  h2("1.  Skill fieldcheck-dp"),
  para([rb("Skill là gì?  "), r("Là bộ hướng dẫn nạp vào Claude, giúp Claude hiểu đúng quy trình xử lý survey của team — từ cách tải data, kiểm tra chất lượng, đến cách tạo bảng.")]),
  sp(80),
  para([rb("Cách cài:")]),
  numbered("Mở Claude Code (ứng dụng desktop)"),
  numbered([r("Vào "), rb("Settings"), r(" → "), rb("Skills")]),
  numbered([r("Nhấn "), rb("Add skill"), r(" → chọn file "), mono("fieldcheck-dp.skill")]),
  numbered("Xong — Claude tự nhận diện khi bạn nhắc đến survey"),
  sp(80),
  callout(["Nếu chưa có file fieldcheck-dp.skill, liên hệ team tech để lấy."], YELLOW_BG, "⚠️"),
  sp(),

  h2("2.  MCP Fieldcheck (kết nối với QMe)"),
  para([rb("MCP là gì?  "), r("Là cầu nối cho phép Claude truy cập thẳng vào hệ thống Fieldcheck / QMe — tự tìm survey, tự tải data mà không cần bạn làm gì thêm.")]),
  sp(80),
  dataTable(
    ["Trạng thái", "Điều xảy ra"],
    [
      [[rb("Không có MCP")], [r("Claude không thể tải data từ QMe — bạn phải tự export file rồi upload")]],
      [[rb("Có MCP", { color: "1A7A1A" })], [r("Chỉ cần nhắn tên survey — Claude tự làm tất cả", { color: "1A7A1A" })]],
    ],
    [3000, 6026]
  ),
  sp(80),
  para([rb("Cách kết nối:")]),
  numbered([r("Mở "), rb("Claude Code"), r(" → "), rb("Settings"), r(" → "), rb("Connections"), r(" (hoặc MCP Servers)")]),
  numbered("Thêm MCP với thông tin do team tech cung cấp"),
  numbered("Đăng nhập tài khoản Fieldcheck khi được yêu cầu"),
  sp(80),
  callout(["Sau khi kết nối, bạn sẽ thấy biểu tượng Fieldcheck trong danh sách kết nối của Claude."], BLUE_LIGHT, "💡"),
  sp(),

  // ─── 3. Bắt đầu ───────────────────────────────────────────────────────────
  h1("Bắt đầu sử dụng"),
  para("Sau khi cài xong, mở Claude Code và nhắn lệnh theo cú pháp:"),
  sp(60),
  codeBlock([
    "/fieldcheck-dp chạy survey VN8966",
    "/fieldcheck-dp làm bảng cho VN8894 - Express",
  ]),
  sp(80),
  para([
    r("Phần "), mono("/fieldcheck-dp"), rb(" là bắt buộc"),
    r(" — đây là cách gọi đúng skill. Nếu thiếu, Claude sẽ không biết dùng quy trình này."),
  ]),
  sp(80),
  callout(
    ["Sau lần đầu khởi động, các tin nhắn tiếp theo trong cùng hội thoại không cần gõ lại /fieldcheck-dp — chỉ cần nhắn bình thường như: \"thêm câu Q8\", \"bật sig test\", \"chạy quality check\"…"],
    GREEN_BG, "💡"
  ),
  sp(80),
  para([rb("Claude hiển thị tiến trình để bạn theo dõi:")]),
  sp(60),
  codeBlock([
    "📋 Pipeline: VN8966",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "✅ 1. Tìm survey       — tìm thấy: VN8966 LIPOVITAN",
    "✅ 2. Tải data         — 450 responses",
    "✅ 3. Xử lý data       — xong",
    "⏳ 4. Chọn bảng       — đang mở giao diện chọn...",
    "⬜ 5. Xuất file Excel",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
  ]),
  sp(),

  // ─── 4. Bước 1-3 ──────────────────────────────────────────────────────────
  h1("Bước 1–3: Claude tự làm"),
  para("Bạn không cần làm gì trong 3 bước đầu. Claude tự tìm survey, tải data và xử lý."),
  sp(80),
  para([rb("Nếu survey đã chạy trước đây"), r(", Claude hỏi:")]),
  callout(
    [[ri("\"Data của VN8966 đã có sẵn (tải ngày 10/06). Dùng data cũ hay tải lại từ QMe?\"")]],
    BLUE_LIGHT, "🤖"
  ),
  sp(80),
  bullet([r("Nhắn "), rb("\"dùng cũ\""), r(" hoặc "), rb("\"ok\""), r(" → bỏ qua bước tải, nhanh hơn")]),
  bullet([r("Nhắn "), rb("\"tải lại\""), r(" → tải data mới nhất từ QMe")]),
  sp(),

  // ─── 5. Quality check ─────────────────────────────────────────────────────
  h1("Bước 3b: Kiểm tra chất lượng data  (tuỳ chọn)"),
  para("Sau khi xử lý data, Claude hỏi:"),
  sp(60),
  callout(
    [[ri("\"Bạn có muốn kiểm tra chất lượng data trước khi tạo bảng không? Tôi sẽ rà soát toàn bộ 450 người xem có ai trả lời sai logic không.\"")]],
    BLUE_LIGHT, "🤖"
  ),
  sp(80),
  para([rb("Nên chạy khi:")]),
  bullet("Survey có nhiều câu điều kiện (câu hiện/ẩn theo câu trả lời trước)"),
  bullet("Cần kiểm tra trước khi giao client"),
  bullet("Muốn biết data sạch hay chưa"),
  sp(80),
  para([r("Nhắn "), rb("\"chạy kiểm tra\""), r(" hoặc "), rb("\"bỏ qua\""), r(" để tiếp tục.")]),
  sp(),

  h2("Kết quả kiểm tra"),
  para("Claude tóm tắt ngay trong chat:"),
  sp(60),
  codeBlock([
    "📊 Kết quả kiểm tra — VN8966",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "Tổng số người     : 450",
    "Có vấn đề         : 12 người (2.7%)",
    "",
    "Loại vấn đề phát hiện:",
    "  ❌  3 người  — bỏ qua câu hỏi bắt buộc",
    "  ⚠️   5 người  — không trả lời câu được dẫn đến",
    "  🔍  2 người  — trả lời câu lẽ ra không được hiển thị",
    "  💥  2 người  — câu trả lời mâu thuẫn nhau",
    "",
    "Câu có vấn đề nhiều nhất:",
    "  Q15 (Mức độ sử dụng) — 6 người",
    "  Q8  (Loại sản phẩm)  — 3 người",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
  ]),
  sp(),

  h2("Ý nghĩa từng loại"),
  dataTable(
    ["Loại vấn đề", "Ý nghĩa", "Nên làm gì"],
    [
      ["Bỏ qua câu bắt buộc",              "Câu luôn phải trả lời nhưng bị bỏ trống",             "Cần xem xét, có thể loại profile"],
      ["Không trả lời câu được dẫn đến",   "Câu đó phải hiện theo logic nhưng không có câu trả lời", "Có thể lỗi survey — báo team"],
      ["Trả lời câu không được hiển thị",  "Câu không nên hiện nhưng vẫn có câu trả lời",          "Lỗi routing QMe — báo team"],
      ["Câu trả lời mâu thuẫn",            "Hai câu trả lời không thể đúng cùng lúc",              "Nên loại trước khi chạy bảng"],
    ],
    [2500, 3300, 3226]
  ),
  sp(),

  h2("Hỏi thêm về kết quả"),
  para("Sau khi xem tóm tắt, bạn có thể hỏi Claude tiếp:"),
  sp(60),
  codeBlock([
    "Xem chi tiết câu Q15",
    "Người nào bị nhiều lỗi nhất?",
    "Chỉ xem những người có câu trả lời mâu thuẫn",
  ]),
  sp(),

  // ─── 6. Datatable Builder ─────────────────────────────────────────────────
  h1("Bước 4: Chọn bảng — Datatable Builder"),
  para([r("Claude mở giao diện "), rb("Datatable Builder"), r(" ngay trong cửa sổ chat.")]),
  sp(),

  h2("Làm theo 4 bước trong Builder"),

  h3("Bước 1 — Cài đặt chung"),
  bullet([r("Đặt tên bảng (ví dụ: "), ri("VN8966 - Data Table"), r(")")]),
  bullet("Chọn loại output: Percentage + Count / Chỉ Percentage / Chỉ Count"),
  sp(80),

  h3("Bước 2 — Banner  (cột tiêu đề ngang)"),
  bullet("Danh sách câu hỏi phân nhóm respondents hiện ra"),
  bullet("Tick chọn câu muốn dùng, ví dụ: Giới tính, Độ tuổi, Khu vực"),
  sp(80),

  h3("Bước 3 — Stub  (câu hỏi theo hàng)"),
  bullet("Danh sách tất cả câu hỏi của survey hiện ra"),
  bullet([r("Tick chọn câu muốn xem kết quả, hoặc tick "), rb("\"Chọn tất cả\"")]),
  sp(80),

  h3("Bước 4 — Xác nhận"),
  bullet("Xem lại toàn bộ cấu hình"),
  bullet([r("Nhấn "), rb("\"Tạo cấu hình bảng\"")]),
  bullet([r("Nhấn "), rb("\"Sao chép\""), r(" → dán vào ô chat → gửi")]),
  sp(80),
  para("Claude nhận được cấu hình, lưu lại và tự chạy."),
  sp(80),
  callout(
    ["Builder chỉ tạo bảng cơ bản. Các tính năng nâng cao như Sig test, Top 2 Box, Mean, Brand header… bạn yêu cầu thêm sau khi đã có file Excel."],
    BLUE_LIGHT, "💡"
  ),
  sp(),

  // ─── 7. Kết quả ───────────────────────────────────────────────────────────
  h1("Bước 5: Nhận file Excel"),
  para("Claude báo khi xong:"),
  sp(60),
  codeBlock([
    "✅ Xong! File Excel đã sẵn sàng.",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "📁 output/VN8966/v1/datatable.xlsx",
    "📊 Sheets: General - Pct | General - Count",
    "👥 450 respondents",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "Bạn muốn thay đổi gì không?",
  ]),
  sp(80),
  para([r("Mở thư mục "), mono("output/VN8966/v1/"), r(" để lấy file.")]),
  sp(),

  // ─── 8. Thay đổi bảng ─────────────────────────────────────────────────────
  h1("Thay đổi bảng sau khi chạy"),
  para([r("Chỉ cần nhắn — không cần mở Builder lại. Claude sửa và tạo file mới "), ri("(v2, v3…)"), r(". "), rb("File cũ vẫn còn nguyên.")]),
  sp(),

  h2("Cột tiêu đề (banner)"),
  dataTable(
    ["Bạn nhắn", "Claude làm"],
    [
      ["\"Thêm câu Income vào cột tiêu đề\"",          "Thêm Income vào banner"],
      ["\"Bỏ cột Độ tuổi\"",                           "Xoá Age khỏi banner"],
      ["\"Thêm cột Total cho từng khu vực\"",           "Thêm tổng trước mỗi nhóm Area"],
      ["\"Dùng Brand làm cột, tất cả brands từ Q17\"",  "Thêm Brand matrix header"],
    ],
    [4600, 4426]
  ),
  sp(),

  h2("Câu hỏi trong bảng (stub) và thống kê"),
  dataTable(
    ["Bạn nhắn", "Claude làm"],
    [
      ["\"Thêm câu Q8 vào bảng\"",             "Thêm Q8"],
      ["\"Bỏ Q15 ra khỏi bảng\"",              "Xoá Q15"],
      ["\"Thêm điểm trung bình cho Q36\"",      "Thêm Mean cho Q36"],
      ["\"Thêm Top 2 Box codes 4 và 5 cho Q8\"","Thêm T2B"],
      ["\"Thêm tất cả câu còn lại\"",           "Thêm hết câu chưa có trong bảng"],
    ],
    [4600, 4426]
  ),
  sp(),

  h2("Loại bảng và định dạng"),
  dataTable(
    ["Bạn nhắn", "Claude làm"],
    [
      ["\"Bật kiểm định thống kê (sig test)\"", "Thêm sheet Sig test"],
      ["\"Tắt sig test\"",                      "Tắt sheet Sig test"],
      ["\"Chỉ giữ sheet Percentage\"",           "Bỏ sheet Count"],
      ["\"Hiển thị 1 chữ số sau dấu phẩy\"",   "Đổi định dạng % thành 35.2%"],
      ["\"Xuất bảng tiếng Anh\"",               "Chạy lại với nhãn tiếng Anh"],
      ["\"Xuất bảng tiếng Việt\"",              "Chạy lại với nhãn tiếng Việt"],
    ],
    [4600, 4426]
  ),
  sp(),

  h2("Cập nhật data"),
  dataTable(
    ["Bạn nhắn", "Claude làm"],
    [
      ["\"Tải data mới nhất từ QMe\"",                  "Fetch lại + chạy lại toàn bộ"],
      ["\"Bao gồm cả người chưa duyệt (pending)\"",     "Thêm pending profiles vào data"],
    ],
    [4600, 4426]
  ),
  sp(),

  // ─── 9. Fallback ──────────────────────────────────────────────────────────
  h1("Khi không tải được data từ QMe"),
  para("Đôi khi kết nối QMe bị treo. Claude sẽ thông báo:"),
  sp(60),
  callout(
    [[ri("\"Tải data bị gián đoạn. Bạn có thể export file từ Fieldcheck và upload vào đây.\"")]],
    BLUE_LIGHT, "🤖"
  ),
  sp(80),
  para([rb("Cách xử lý:")]),
  numbered([r("Mở "), rb("Fieldcheck"), r(" → tìm survey → "), rb("Export"), r(" → tải file "), mono(".zip"), r(" về máy")]),
  numbered([r("Kéo thả file "), mono(".zip"), r(" vào cửa sổ chat")]),
  numbered("Claude tự xử lý và tiếp tục"),
  sp(),

  // ─── 10. Câu lệnh nhanh ───────────────────────────────────────────────────
  h1("Câu nói nhanh hay dùng"),

  h2("Lần đầu trong hội thoại — cần có /fieldcheck-dp ở đầu"),
  codeBlock([
    "/fieldcheck-dp chạy survey VN8966",
    "/fieldcheck-dp làm bảng cho VN8894 - Express",
    "/fieldcheck-dp chạy quality check VN8966",
  ]),
  sp(),

  h2("Các lần tiếp theo — nhắn bình thường"),
  codeBlock([
    "chạy kiểm tra chất lượng data",
    "xem chi tiết câu Q15",
    "người nào bị nhiều lỗi nhất?",
    "thêm [tên câu] vào cột tiêu đề",
    "thêm [tên câu] vào bảng",
    "bỏ [tên câu] khỏi bảng",
    "thêm điểm trung bình cho [câu]",
    "thêm Top 2 Box cho [câu]",
    "bật / tắt sig test",
    "chỉ giữ sheet Percentage",
    "xuất tiếng Anh / xuất tiếng Việt",
    "tải lại data mới",
    "file kết quả ở đâu?",
  ]),
  sp(),

  // ─── 11. File output ──────────────────────────────────────────────────────
  h1("File output ở đâu?"),
  para([r("Tất cả file nằm trong thư mục "), mono("output/"), r(" trong cùng thư mục với Claude:")]),
  sp(60),
  codeBlock([
    "output/",
    "└── VN8966/",
    "    ├── v1/datatable.xlsx         ← Bảng lần đầu",
    "    ├── v2/datatable.xlsx         ← Sau lần chỉnh đầu tiên",
    "    ├── v3/datatable.xlsx         ← ...",
    "    └── quality/",
    "        └── flagged_profiles.csv  ← Danh sách người có vấn đề",
  ]),
  sp(80),
  para([r("Mỗi lần thay đổi bảng → Claude tạo thêm thư mục mới (v2, v3…). "), rb("Các file cũ không bị xoá.")]),
  sp(),

  // ─── 12. FAQ ──────────────────────────────────────────────────────────────
  h1("Câu hỏi thường gặp"),

  ...faq("Tôi không thấy Builder hiện ra, phải làm gì?",
    [r("Thử nhắn lại: "), rb("\"mở Datatable Builder\""), r(" hoặc "), rb("\"cho tôi chọn cột và hàng\""), r(".")]),

  ...faq("File Excel nằm ở đâu?",
    [r("Nhắn Claude: "), rb("\"file ở đâu?\""), r(" — Claude sẽ chỉ đường dẫn chính xác.")]),

  ...faq("Thay đổi bảng có mất file cũ không?",
    [r("Không. Mỗi lần thay đổi Claude tạo v2, v3… file v1 vẫn còn nguyên.")]),

  ...faq("Kiểm tra chất lượng data có bắt buộc không?",
    [r("Không. Nhưng nên chạy nếu survey có nhiều câu hiện/ẩn theo điều kiện, hoặc trước khi giao file cho client.")]),

  ...faq("Bảng ra thiếu người / 0 người?",
    [r("Nhắn: "), rb("\"bao gồm cả người chưa duyệt (pending)\""), r(" — có thể data chỉ có profile pending chưa được approve.")]),

  ...faq("Tải data từ QMe bị lỗi?",
    [r("Export file zip từ Fieldcheck rồi kéo vào chat — Claude sẽ nhận và tiếp tục tự động.")]),

  ...faq("Claude có tự sửa file mà không hỏi không?",
    [r("Không. Claude luôn báo trước khi tải lại data hoặc ghi đè file quan trọng. Với thay đổi bảng thông thường, Claude sửa thẳng và thông báo kết quả.")]),

  ...faq("Muốn bảng tiếng Anh?",
    [r("Nhắn: "), rb("\"xuất bảng tiếng Anh\""), r(" bất cứ lúc nào — Claude chạy lại với nhãn tiếng Anh.")]),
];

// ── Assemble & save ───────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: DXA(0.35), hanging: DXA(0.2) } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: DXA(0.4), hanging: DXA(0.25) } } },
        }],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22, color: "222222" } },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: WHITE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: BLUE_MID },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: DXA(16.54) },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          children: [new TextRun({ text: "SurveyFlow + FieldCheck — Hướng dẫn sử dụng", font: "Arial", size: 18, color: "888888" })],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
          spacing: { after: 0 },
        }),
      ]}),
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          children: [
            new TextRun({ text: "Trang ", font: "Arial", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "888888" }),
            new TextRun({ text: " / ", font: "Arial", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: "888888" }),
          ],
          alignment: AlignmentType.RIGHT,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
          spacing: { before: 80, after: 0 },
        }),
      ]}),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("docs/CLAUDE_CHAT_GUIDE.docx", buf);
  console.log("✅  Saved: docs/CLAUDE_CHAT_GUIDE.docx");
});
