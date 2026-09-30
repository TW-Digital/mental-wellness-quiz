/**
 * Mental Wellness Check-In — Google Apps Script backend
 *
 * บันทึกเฉพาะข้อมูลกิจกรรมลงชีต "Responses":
 *   Timestamp | EmployeeID | Status | TotalScore | OverallMode | SessionID
 *
 * - กด "เริ่มตอบคำถาม"  → เพิ่มแถวใหม่ Status = Started
 * - กด "ดูผลลัพธ์ของฉัน" → อัปเดตแถวเดิม (ค้นจาก SessionID) เป็น Completed + TotalScore + OverallMode
 *                          และเปลี่ยน Timestamp เป็นเวลาที่ส่งแบบทดสอบ
 */

const SHEET_NAME = 'Responses';
const HEADERS = ['Timestamp', 'EmployeeID', 'Status', 'TotalScore', 'OverallMode', 'SessionID'];
const TZ = 'Asia/Bangkok';

// เปิดหน้าเว็บ (กรณีเสิร์ฟ index.html ผ่าน Apps Script โดยตรง)
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Mental Wellness Check-In')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// รับข้อมูลจากหน้าเว็บที่โฮสต์ที่อื่น (fetch POST)
function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const result = logEvent(payload);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// เรียกได้ทั้งจาก doPost และ google.script.run
function logEvent(p) {
  const employeeId = String(p.employeeId || '');
  if (!/^\d{8}$/.test(employeeId)) return { ok: false, error: 'invalid employeeId' };
  const sessionId = String(p.sessionId || '').slice(0, 40);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

    if (p.action === 'start') {
      sheet.appendRow([now, "'" + employeeId, 'Started', '', '', sessionId]);
      return { ok: true };
    }

    if (p.action === 'submit') {
      const total = Math.max(0, Math.min(96, Number(p.totalScore) || 0));
      const mode = String(p.overallMode || '').slice(0, 60);
      const row = findRow_(sheet, sessionId);
      if (row) {
        sheet.getRange(row, 1, 1, 5).setValues([[now, "'" + employeeId, 'Completed', total, mode]]);
      } else {
        sheet.appendRow([now, "'" + employeeId, 'Completed', total, mode, sessionId]);
      }
      return { ok: true };
    }

    return { ok: false, error: 'unknown action' };
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRow_(sheet, sessionId) {
  if (!sessionId) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const ids = sheet.getRange(2, 6, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === sessionId) return i + 2;
  }
  return null;
}
