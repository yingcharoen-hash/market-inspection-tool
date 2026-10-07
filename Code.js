/**
 * Apps Script สำหรับรับข้อมูลจากเครื่องมือตรวจตลาดสด แล้วบันทึกลง Google Sheet
 * + อัปโหลดรูปภาพไป Google Drive อัตโนมัติ
 *
 * คอลัมน์ใน Sheet1 (เรียงตามลำดับ):
 * วันที่ | ผู้ตรวจ | หมวด | สถานที่/โซน | รายการ | ผลตรวจ | ประเภทการรุกล้ำ | หมายเหตุ | รูปภาพ1 | รูปภาพ2 | เวลาบันทึก
 */

const SPREADSHEET_ID = "1BGtG2oCqu6yTBm_cqE71ReRhKVjrGsqoMmIlKuPx6qc";

function doPost(e) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const data = JSON.parse(e.postData.contents);
  const folder = getOrCreateFolder("ภาพตรวจตลาดสด");

  const lastRow = sheet.getLastRow();
  let existingData = [];
  if (lastRow > 1) {
    const startRow = Math.max(2, lastRow - 200);
    const numRows = lastRow - startRow + 1;
    existingData = sheet.getRange(startRow, 1, numRows, 11).getValues(); 
  }

  const parseTHDate = (str) => {
    if(!str) return 0;
    try {
      const parts = String(str).split(' ');
      const dParts = parts[0].split('/');
      if (dParts.length !== 3) return 0;
      const iso = `${dParts[2]}-${dParts[1]}-${dParts[0]}T${parts[1]||'00:00:00'}`;
      return new Date(iso).getTime();
    } catch(e) { return 0; }
  };

  const isDuplicate = (r) => {
    const rTime = parseTHDate(r.timestamp);
    if (!rTime) return false;
    
    return existingData.some(existing => {
      const eTime = parseTHDate(existing[10]);
      if (!eTime) return false;
      const diffMs = Math.abs(rTime - eTime);
      
      return diffMs <= 5 * 60 * 1000 && // Within 5 minutes
        String(existing[0]) === String(r.date) &&
        String(existing[1]) === String(r.guardName || "") &&
        String(existing[2]) === String(r.section) &&
        String(existing[3]) === String(r.location) &&
        String(existing[4]) === String(r.item) &&
        String(existing[5]) === String(r.status) &&
        String(existing[6]) === String(r.encType || "") &&
        String(existing[7]) === String(r.shopName || "") &&
        String(existing[11] || "") === String(r.note || "");
    });
  };

  data.rows.forEach(row => {
    if (isDuplicate(row)) {
      return; // Skip duplicate double submission
    }

    const links = (row.photos || []).map(p => {
      if (!p.base64) return "";
      return saveImage(p.base64, p.name, folder);
    }).filter(x => x);

    const newRow = [
      row.date, row.guardName || "", row.section, row.location,
      row.item, row.status, row.encType || "",
      row.shopName || "", links[0] || "", links[1] || "", row.timestamp, row.note || ""
    ];
    
    sheet.appendRow(newRow);
    existingData.push([...newRow.slice(0, 5), null, null, null, null, null, newRow[10]]);
  });

  // ล้าง cache ทั้งหมดเมื่อมีข้อมูลใหม่
  const cache = CacheService.getScriptCache();
  cache.removeAll(['data_90d_chunks','data_all_chunks',
    'data_90d_0','data_90d_1','data_90d_2','data_90d_3','data_90d_4',
    'data_all_0','data_all_1','data_all_2','data_all_3','data_all_4']);

  return ContentService.createTextOutput("OK");
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── action=data: ส่งข้อมูลให้ dashboard (chunked cache + กรอง 90 วัน) ──
  if (action === 'data') {
    const allHistory = (e.parameter.days === 'all');
    const cacheKey = allHistory ? 'data_all' : 'data_90d';
    const cache = CacheService.getScriptCache();

    // โหลดจาก chunked cache
    const chunkCount = parseInt(cache.get(cacheKey + '_chunks') || '0');
    if (chunkCount > 0) {
      const keys = Array.from({length: chunkCount}, (_, i) => cacheKey + '_' + i);
      const chunks = cache.getAll(keys);
      const combined = keys.map(k => chunks[k] || '').join('');
      if (combined) return ContentService.createTextOutput(combined).setMimeType(ContentService.MimeType.JSON);
    }

    // ดึงจาก Sheet
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
    const dataRows = sheet.getRange(2,1,lastRow-1,lastCol).getValues();

    // กรองเฉพาะ 90 วันล่าสุด (default)
    const cutoff = allHistory ? null : new Date(Date.now() - 90*24*60*60*1000);
    const rows = dataRows.map(row => {
      const obj = {};
      headers.forEach((h,i) => { let v=row[i]; if(v instanceof Date) v=v.toISOString(); obj[h]=v; });
      return obj;
    }).filter(obj => {
      if (!cutoff) return true;
      const d = new Date(obj[headers[0]]);
      return isNaN(d) || d >= cutoff;
    });

    const result = JSON.stringify({rows});

    // บันทึก chunked cache (90KB ต่อ chunk)
    const CHUNK = 90000;
    const n = Math.ceil(result.length / CHUNK);
    const cacheObj = {[cacheKey+'_chunks']: String(n)};
    for (let i=0;i<n;i++) cacheObj[cacheKey+'_'+i] = result.slice(i*CHUNK,(i+1)*CHUNK);
    try { cache.putAll(cacheObj, 300); } catch(_) {}

    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  }

  // ── action=shops ──
  if (action === 'shops') {
    try {
      const SHOP_SS_ID = "1VqB8yiqny-UZNbY12AJZNecYPDff0-qkrv9kJmy-bfQ";
      const SHOP_GID = 1336386039;
      const extSheet = SpreadsheetApp.openById(SHOP_SS_ID).getSheets().find(s=>s.getSheetId()===SHOP_GID);
      let shops = [];
      if (extSheet && extSheet.getLastRow()>=2) {
        shops = extSheet.getRange(2,11,extSheet.getLastRow()-1,9).getValues()
          .filter(r=>String(r[0]).trim()==='Active').map(r=>String(r[8]).trim()).filter(v=>v!=="");
        shops = [...new Set(shops)].sort((a,b)=>a.localeCompare(b,'th'));
      }
      return ContentService.createTextOutput(JSON.stringify({shops})).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({shops:[],error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── action=shops2: ดึงชื่อร้านค้าพร้อมโซนจาก spreadsheet ทะเบียนร้านค้า ──
  if (action === 'shops2') {
    try {
      const SHOP_SS_ID = "1VqB8yiqny-UZNbY12AJZNecYPDff0-qkrv9kJmy-bfQ";
      const SHOP_GID = 1336386039;
      const extSheet = SpreadsheetApp.openById(SHOP_SS_ID).getSheets().find(s=>s.getSheetId()===SHOP_GID);
      let shops = [];
      if (extSheet && extSheet.getLastRow()>=2) {
        let shopList = extSheet.getRange(2, 6, extSheet.getLastRow() - 1, 14).getValues()
          .filter(r => String(r[5]).trim() === 'Active')              
          .map(r => ({ zone: String(r[0]).trim(), name: String(r[13]).trim() })) 
          .filter(v => v.name !== "");
          
        const seen = new Set();
        shopList.forEach(item => {
          const key = item.name + '|' + item.zone;
          if (!seen.has(key)) {
            seen.add(key);
            shops.push(item);
          }
        });
        shops.sort((a, b) => a.name.localeCompare(b.name, 'th'));
      }
      return ContentService.createTextOutput(JSON.stringify({ shops }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ shops: [], error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── action=admin_data: ส่งข้อมูลเฉพาะสำหรับฝ่ายขาย (อ่านแค่ 3000 บรรทัดล่าสุดให้เร็วที่สุด) ──
  if (action === 'admin_data') {
    try {
      const sheet = ss.getSheets()[0];
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);

      // อ่านแค่ 3000 บรรทัดล่าสุด
      const startRow = Math.max(2, lastRow - 3000);
      const numRows = lastRow - startRow + 1;

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const dataRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

      const rows = [];
      const statusIdx = headers.indexOf('ผลตรวจ');
      const penaltyIdx = headers.indexOf('สถานะบทลงโทษ');

      for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const statusVal = statusIdx > -1 ? row[statusIdx] : '';
        if (statusVal === 'พบการรุกล้ำ' || statusVal === 'ไม่ผ่าน') {
           const obj = {};
           headers.forEach((h, colIdx) => { 
             let v = row[colIdx]; 
             if(v instanceof Date) v = v.toISOString(); 
             obj[h] = v; 
           });
           rows.push(obj);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ rows: rows })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── action=admin: โหลดหน้าเว็บสำหรับฝ่ายขาย (Web App) ──
  if (action === 'admin') {
    const email = Session.getActiveUser().getEmail();
    const adminSheet = ss.getSheetByName("Admin_Users");
    let isAdmin = false;
    
    if (adminSheet) {
      const lastR = adminSheet.getLastRow();
      if (lastR > 1) {
        const admins = adminSheet.getRange(2, 1, lastR - 1, 1).getValues().map(r => String(r[0]).trim());
        isAdmin = admins.includes(email);
      }
    } else {
      isAdmin = true; 
    }

    if (!isAdmin) {
      return HtmlService.createHtmlOutput('<h2>Access Denied</h2><p>อีเมล ' + email + ' ไม่มีสิทธิ์เข้าถึงหน้านี้ โปรดแจ้งผู้ดูแลระบบ</p>');
    }
    
    const tpl = HtmlService.createTemplateFromFile('SalesAdmin');
    try {
      tpl.initialData = getPendingActions();
    } catch(err) {
      tpl.initialData = JSON.stringify({ error: err.message });
    }
    return tpl.evaluate()
      .setTitle('ระบบจัดการบทลงโทษ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── default: รายชื่อผู้ตรวจ ──
  const sheet = ss.getSheetByName("setting");
  let names = [];
  if (sheet && sheet.getLastRow()>1) {
    names = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues()
      .map(r=>String(r[0]).trim()).filter(v=>v!=="");
  }
  return ContentService.createTextOutput(JSON.stringify({names})).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder(name) {
  const f = DriveApp.getFoldersByName(name);
  return f.hasNext() ? f.next() : DriveApp.createFolder(name);
}

// ── ฟังก์ชันสำหรับ Web App ฝ่ายขาย ──

// ดึงข้อมูลสำหรับตารางฝ่ายขาย
function getPendingActions(offset = 0, limit = 500) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return JSON.stringify({rows:[], hasMore: false, total: 0});

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    // อ่านทั้งหมดเลย เพราะ getValues เร็วมาก (ปัญหาอยู่ที่ตอนส่ง JSON กลับ)
    const dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const statusIdx = headers.indexOf('ผลตรวจ');
    const penaltyIdx = headers.indexOf('สถานะบทลงโทษ');

    const allPending = [];
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const statusVal = statusIdx > -1 ? row[statusIdx] : '';
      const penaltyVal = penaltyIdx > -1 ? row[penaltyIdx] : '';
      
      if (statusVal === 'พบการรุกล้ำ' || statusVal === 'ไม่ผ่าน') {
         allPending.push(row);
      }
    }

    // ตัดแบ่งข้อมูลตาม offset และ limit
    const slice = allPending.slice(offset, offset + limit);
    const rows = [];
    for (const row of slice) {
       const obj = {};
       headers.forEach((h, colIdx) => { 
         let v = row[colIdx]; 
         if(v instanceof Date) v = v.toISOString(); 
         obj[h] = v; 
       });
       rows.push(obj);
    }
    
    return JSON.stringify({ 
      rows: rows, 
      hasMore: (offset + limit) < allPending.length,
      total: allPending.length
    });
  } catch (err) {
    return JSON.stringify({ error: err.message, stack: err.stack });
  }
}

// รับข้อมูลการอัปเดตบทลงโทษจากหน้าเว็บ
function updateResolutionData(payload) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  
  const timestampToFind = String(payload.targetTimestamp);
  let targetRow = -1;
  let tsIndex = headers.indexOf('เวลาบันทึก');
  if (tsIndex === -1) tsIndex = 10;

  for (let i = 1; i < allData.length; i++) {
    let rowTs = allData[i][tsIndex];
    if (rowTs instanceof Date) rowTs = rowTs.toISOString();
    if (String(rowTs) === timestampToFind) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow > -1) {
    let actionCol = headers.indexOf('สถานะบทลงโทษ');
    if (actionCol === -1) {
      actionCol = sheet.getLastColumn(); 
      // สร้าง Header ใหม่ถ้ายังไม่มี
      sheet.getRange(1, actionCol + 1).setValue('สถานะบทลงโทษ');
      sheet.getRange(1, actionCol + 2).setValue('รายละเอียด');
      sheet.getRange(1, actionCol + 3).setValue('ผู้ดำเนินการ');
      sheet.getRange(1, actionCol + 4).setValue('วันที่ดำเนินการ');
      sheet.getRange(1, actionCol + 5).setValue('ไฟล์แนบหลักฐาน');
    }
    
    let linkUrl = "";
    if (payload.photoBase64) {
      const folder = getOrCreateFolder("ภาพตรวจตลาดสด");
      linkUrl = saveImage(payload.photoBase64, "Evidence_" + Date.now(), folder) || "";
    }

    const email = Session.getActiveUser().getEmail();
    
    sheet.getRange(targetRow, actionCol + 1).setValue(payload.penaltyStep);
    sheet.getRange(targetRow, actionCol + 2).setValue(payload.note);
    sheet.getRange(targetRow, actionCol + 3).setValue(email);
    sheet.getRange(targetRow, actionCol + 4).setValue(new Date());
    if (linkUrl) {
      sheet.getRange(targetRow, actionCol + 5).setValue(linkUrl);
    }
    
    const cache = CacheService.getScriptCache();
    cache.removeAll(['data_90d_chunks','data_all_chunks']);
    
    return { success: true };
  }
  return { success: false, error: 'ไม่พบรายการที่ต้องการอัปเดต' };
}

function saveImage(base64, filename, folder) {
  const parts = base64.split(",");
  const bytes = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(bytes,"image/jpeg",filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ─────────────────────────────────────────────
// warmCache: เรียกทุก 5 นาทีผ่าน Time trigger เพื่อป้องกัน cold start
// ─────────────────────────────────────────────
function warmCache() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const dataRows = sheet.getRange(2,1,lastRow-1,lastCol).getValues();

  const cutoff = new Date(Date.now() - 90*24*60*60*1000);
  const rows = dataRows.map(row => {
    const obj = {};
    headers.forEach((h,i) => { let v=row[i]; if(v instanceof Date) v=v.toISOString(); obj[h]=v; });
    return obj;
  }).filter(obj => {
    const d = new Date(obj[headers[0]]);
    return isNaN(d) || d >= cutoff;
  });

  const result = JSON.stringify({rows});
  const CHUNK = 90000;
  const n = Math.ceil(result.length / CHUNK);
  const cacheObj = {'data_90d_chunks': String(n)};
  for (let i=0;i<n;i++) cacheObj['data_90d_'+i] = result.slice(i*CHUNK,(i+1)*CHUNK);
  try { CacheService.getScriptCache().putAll(cacheObj, 360); } catch(_) {}
}

function setupTrigger() {
  // ลบ trigger เก่าของ warmCache ก่อน
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'warmCache')
    .forEach(t => ScriptApp.deleteTrigger(t));
  // สร้าง trigger ใหม่ทุก 5 นาที
  ScriptApp.newTrigger('warmCache')
    .timeBased()
    .everyMinutes(5)
    .create();
}
function moveDuplicatesToDuplicateSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mainSheet = ss.getSheets()[0];
  let sheetDuplicate = ss.getSheetByName("ข้อมูลซ้ำ");
  
  if (!sheetDuplicate) {
    sheetDuplicate = ss.insertSheet("ข้อมูลซ้ำ");
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues();
    sheetDuplicate.appendRow(headers[0]);
  }
  
  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();
  if (lastRow <= 1) return;
  
  const data = mainSheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  
  const parseTHDate = (str) => {
    if(!str) return 0;
    try {
      const parts = String(str).split(' ');
      const dParts = parts[0].split('/');
      if (dParts.length !== 3) return 0;
      const iso = `${dParts[2]}-${dParts[1]}-${dParts[0]}T${parts[1]||'00:00:00'}`;
      return new Date(iso).getTime();
    } catch(e) { return 0; }
  };

  const seenMap = new Map(); 
  const uniqueData = [];
  const duplicateData = []; 
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Strict Key: Date(0) + Guard(1) + Section(2) + Location(3) + Item(4) + Status(5) + EncType(6) + ShopName(7) + Note(11)
    const key = String(row[0]).trim() + "|" + 
                String(row[1]).trim() + "|" + 
                String(row[2]).trim() + "|" + 
                String(row[3]).trim() + "|" + 
                String(row[4]).trim() + "|" + 
                String(row[5]).trim() + "|" + 
                String(row[6]).trim() + "|" + 
                String(row[7]).trim() + "|" + 
                String(row[11]).trim();
                
    const rTime = parseTHDate(row[10]);
    
    if (seenMap.has(key)) {
      const eTime = seenMap.get(key);
      let isDuplicate = false;
      
      // If timestamps are valid, check if within 5 mins
      if (rTime > 0 && eTime > 0) {
        if (Math.abs(rTime - eTime) <= 5 * 60 * 1000) {
          isDuplicate = true;
        }
      } else {
        // If timestamps are missing but the strict key matches exactly, treat as duplicate
        isDuplicate = true;
      }
      
      if (isDuplicate) {
         duplicateData.push(row);
      } else {
         seenMap.set(key, rTime);
         uniqueData.push(row);
      }
    } else {
      seenMap.set(key, rTime);
      uniqueData.push(row);
    }
  }
  
  if (duplicateData.length > 0) {
    sheetDuplicate.getRange(sheetDuplicate.getLastRow() + 1, 1, duplicateData.length, lastCol).setValues(duplicateData);
  }
  
  mainSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (uniqueData.length > 0) {
    mainSheet.getRange(2, 1, uniqueData.length, lastCol).setValues(uniqueData);
  }
}
