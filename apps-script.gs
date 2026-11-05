// Dán toàn bộ file này vào Google Sheet: Tiện ích mở rộng → Apps Script.
// Nhận điểm từ trang chơi (doPost) và trả danh sách điểm cho trang xếp hạng (doGet).

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 30 người nộp cùng lúc không đè dòng nhau
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(["Thời điểm nộp", "Tên", "Số câu đúng", "Tổng thời gian (giây)", "Chi tiết"]);
    const d = JSON.parse(e.postData.contents);
    // tên bắt đầu bằng = + - @ sẽ bị Sheet chạy như công thức, thêm dấu ' để giữ là chữ
    const name = String(d.name || "").trim().slice(0, 40).replace(/^[=+\-@]/, "'$&");
    const score = Math.max(0, Math.min(20, Math.floor(Number(d.score) || 0)));
    const sec = Math.max(0, Number(d.ms) || 0) / 1000;
    if (!name) return out({ ok: false, error: "Thiếu tên" });
    sheet.appendRow([new Date(), name, score, Math.round(sec * 100) / 100, String(d.detail || "").slice(0, 500)]);
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getDataRange().getValues().slice(1);
  return out(rows.filter(r => r[1] !== "").map(r => ({ t: new Date(r[0]).getTime(), name: r[1], score: r[2], sec: r[3] })));
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
