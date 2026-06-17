/**
 * Apps Script สำหรับรับข้อมูลจากเครื่องมือตรวจตลาดสด แล้วบันทึกลง Google Sheet
 * + อัปโหลดรูปภาพไป Google Drive อัตโนมัติ
 *
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet ที่ต้องการบันทึกข้อมูล
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดเดิมในไฟล์ Code.gs ทั้งหมด แล้ววางโค้ดนี้แทน
 * 4. Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. กด Deploy แล้วอนุญาตสิทธิ์ที่ขอ (Sheet + Drive)
 * 6. คัดลอกลิงก์ที่ลงท้ายด้วย /exec แล้วนำไปวางในตัวแปร SHEET_URL
 *    ของไฟล์ market_inspection_tool.html (ค้นหา "SHEET_URL")
 *
 * คอลัมน์ใน Sheet (เรียงตามลำดับ):
 * วันที่ | ผู้ตรวจ | หมวด | สถานที่/โซน | รายการ | ผลตรวจ | หมายเหตุ | รูปภาพ1 | รูปภาพ2 | เวลาบันทึก
 *
 * ชีต "setting": สร้างชีตใหม่ชื่อ "setting" คอลัมน์ A แถวที่ 2 เป็นต้นไป
 * ใส่รายชื่อ รปภ. ทีละแถว — เครื่องมือจะดึงรายชื่อนี้ไปแสดงเป็น dropdown อัตโนมัติ
 *
 * รายละเอียดเพิ่มเติม: ดูไฟล์ google_sheet_setup.md
 */

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  return ContentService.createTextOutput(JSON.stringify({ names: names }))
    .setMimeType(ContentService.MimeType.JSON);
}

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
