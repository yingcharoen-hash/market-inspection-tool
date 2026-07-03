/**
 * Apps Script สำหรับรับข้อมูลจากเครื่องมือตรวจตลาดสด แล้วบันทึกลง Google Sheet
 * + อัปโหลดรูปภาพไป Google Drive อัตโนมัติ
 *
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet ที่ต้องการบันทึกข้อมูล
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดเดิมในไฟล์ Code.gs ทั้งหมด แล้ววางโค้ดนี้แทน
 * 4. Deploy > Manage deployments > แก้ไข deployment ปัจจุบัน > New version > Deploy
 *
 * คอลัมน์ใน Sheet1 (เรียงตามลำดับ):
 * วันที่ | ผู้ตรวจ | หมวด | สถานที่/โซน | รายการ | ผลตรวจ | หมายเหตุ | รูปภาพ1 | รูปภาพ2 | เวลาบันทึก
 *
 * ชีต "setting": สร้างชีตใหม่ชื่อ "setting" คอลัมน์ A แถวที่ 2 เป็นต้นไป
 * ใส่รายชื่อ รปภ. ทีละแถว — เครื่องมือจะดึงรายชื่อนี้ไปแสดงเป็น dropdown อัตโนมัติ
 */

// ID ของ Google Sheet ที่บันทึกข้อมูลตรวจ
// (ดูจาก URL: https://docs.google.com/spreadsheets/d/【ID】/edit)
const SPREADSHEET_ID = "1BGtG2oCqu6yTBm_cqE71ReRhKVjrGsqoMmIlKuPx6qc";

// ─────────────────────────────────────────────
// รับข้อมูลตรวจจากแอป และบันทึกลง Sheet1
// ─────────────────────────────────────────────
function doPost(e) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const data = JSON.parse(e.postData.contents);
  const folder = getOrCreateFolder("ภาพตรวจตลาดสด");

  data.rows.forEach(row => {
    const links = (row.photos || []).map(p => {
      if (!p.base64) return "";
      return saveImage(p.base64, p.name, folder);
    }).filter(x => x);

    sheet.appendRow([
      row.date,
      row.guardName || "",
      row.section,
      row.location,
      row.item,
      row.status,
      row.note,
      links[0] || "",
      links[1] || "",
      row.timestamp
    ]);
  });

  return ContentService.createTextOutput("OK");
}

// ─────────────────────────────────────────────
// doGet: รองรับ 3 endpoint
//   ?action=data   → คืนข้อมูลทุกแถวใน Sheet1 (สำหรับ dashboard)
//   ?action=shops  → คืนรายชื่อร้านค้าจาก spreadsheet ทะเบียน
//   (ไม่มี action) → คืนรายชื่อผู้ตรวจจากชีต setting
// ─────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── action=data: ส่งข้อมูลทั้งหมดให้ dashboard ──
  if (action === 'data') {
    const sheet = ss.getSheets()[0]; // Sheet1 (แถวแรกเป็น header)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return ContentService.createTextOutput(JSON.stringify({ rows: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rows = dataRows.map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        // แปลง Date object → ISO string เพื่อให้ dashboard แสดงวันที่ถูกต้อง
        if (val instanceof Date) val = val.toISOString();
        obj[h] = val;
      });
      return obj;
    });
    return ContentService.createTextOutput(JSON.stringify({ rows }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── action=shops: ดึงชื่อร้านค้าจาก spreadsheet ทะเบียนร้านค้า ──
  if (action === 'shops') {
    try {
      const SHOP_SS_ID = "1VqB8yiqny-UZNbY12AJZNecYPDff0-qkrv9kJmy-bfQ";
      const SHOP_GID   = 1336386039;
      const extSS    = SpreadsheetApp.openById(SHOP_SS_ID);
      const extSheet = extSS.getSheets().find(s => s.getSheetId() === SHOP_GID);
      let shops = [];
      if (extSheet) {
        const lastRow = extSheet.getLastRow();
        if (lastRow >= 2) {
          shops = extSheet.getRange(2, 19, lastRow - 1, 1).getValues()
            .map(r => String(r[0]).trim())
            .filter(v => v !== "");
          shops = [...new Set(shops)].sort((a, b) => a.localeCompare(b, 'th'));
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ shops }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ shops: [], error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── default: รายชื่อผู้ตรวจจากชีต setting ──
  const sheet = ss.getSheetByName("setting");
  let names = [];
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      names = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        .map(r => String(r[0]).trim())
        .filter(v => v !== "");
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ names }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// Helper functions
// ─────────────────────────────────────────────
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function saveImage(base64, filename, folder) {
  const parts = base64.split(",");
  const bytes = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(bytes, "image/jpeg", filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
