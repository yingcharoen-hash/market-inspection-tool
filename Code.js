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

  data.rows.forEach(row => {
    const links = (row.photos || []).map(p => {
      if (!p.base64) return "";
      return saveImage(p.base64, p.name, folder);
    }).filter(x => x);

    sheet.appendRow([
      row.date, row.guardName || "", row.section, row.location,
      row.item, row.status, row.encType || "",
      row.shopName || "", links[0] || "", links[1] || "", row.timestamp, row.note || ""
    ]);
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