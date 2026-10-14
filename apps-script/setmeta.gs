/* ============================================================
   reel2map 網頁寫回「收藏」「我的說明」用的 action=setmeta
   安裝方式：
   1. Apps Script 編輯器左邊「檔案 +」→ 指令碼，命名 setmeta，貼上這整個檔案
   2. 在原本 Code.gs 的 doGet(e) 第一行（token 檢查之前）加上：
        if (e.parameter.action === 'setmeta') return setMeta_(e);
   3. 管理部署作業 → 既有部署的鉛筆 → 版本選「新版本」→ 部署
      （不要按「新增部署作業」，會換網址）

   NOTE_KEY 會寫在公開的 index.html 裡，所以它只能改這兩欄，
   不能新增/刪除地點；原本的密語不需要也不應該放進網頁。
   這個檔案不依賴 Code.gs 裡的任何常數或函式，名稱也都加了 NOTE_ 前綴避免撞名。
   ============================================================ */

const NOTE_KEY = 'r2m-note-7Kq3vX9p';
// 每張地圖對應的試算表與分頁（gid = 網頁讀的那個發布分頁）；網頁用 who=jun 指定，沒帶就是 lolo
const NOTE_SHEETS = {
  lolo: { id: '1PZ3EsCNlXxmh4wL05wmchivjNAJTVS9bW0W3jD4TUkw', gid: 46533551 },     // reel2map_地點資料
  jun:  { id: '156cfCE8yQ5xQJSX5i6g8bJRxCmARRWFFc4WO0tELvH4', gid: 1634312988 },   // reel2map_Jun
};

function noteSheet_(who) {
  const cfg = NOTE_SHEETS.hasOwnProperty(who) ? NOTE_SHEETS[who] : NOTE_SHEETS.lolo;
  const ss = SpreadsheetApp.openById(cfg.id);
  return ss.getSheets().find(s => s.getSheetId() === cfg.gid) || ss.getSheets()[0];
}

function setMeta_(e) {
  const p = e.parameter;
  if (p.key !== NOTE_KEY) return noteText_('forbidden');
  if (p.op === 'read') return readMeta_(p.who);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = noteSheet_(p.who);
    const data = sheet.getDataRange().getValues();
    const header = data[0].map(h => String(h).trim());
    const iName = header.indexOf('名稱'), iLat = header.indexOf('緯度'), iLng = header.indexOf('經度');
    const name = String(p.name || '').trim();
    const lat = parseFloat(p.lat), lng = parseFloat(p.lng);

    // 用「名稱＋座標」找列；座標容許 0.0001 度誤差（試算表顯示位數可能比實際少）
    let row = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][iName]).trim() === name &&
          Math.abs(parseFloat(data[r][iLat]) - lat) < 1e-4 &&
          Math.abs(parseFloat(data[r][iLng]) - lng) < 1e-4) { row = r + 1; break; }
    }
    if (row < 0) return noteText_('not found');

    // 欄位不存在就在最右邊自動補上標題
    const col = title => {
      let i = header.indexOf(title);
      if (i < 0) { i = header.length; header.push(title); sheet.getRange(1, i + 1).setValue(title); }
      return i + 1;
    };
    if ('fav' in p) sheet.getRange(row, col('收藏')).setValue(p.fav === '1' ? '♥' : '');
    if ('memo' in p) {
      const cell = sheet.getRange(row, col('我的說明'));
      cell.setNumberFormat('@');   // 純文字，避免「3/5」被轉成日期
      cell.setValue(noteSanitize_(p.memo));
    }
    return noteText_('ok');
  } finally {
    lock.releaseLock();
  }
}

// op=read：即時讀出有收藏或說明的列（發布的 CSV 會快取 5 分鐘，這個不會）
function readMeta_(who) {
  const data = noteSheet_(who).getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const iName = header.indexOf('名稱'), iLat = header.indexOf('緯度'), iLng = header.indexOf('經度'),
        iFav = header.indexOf('收藏'), iMemo = header.indexOf('我的說明');
  const out = [];
  for (let r = 1; r < data.length; r++) {
    const fav = iFav >= 0 && String(data[r][iFav]).trim() !== '';
    const memo = iMemo >= 0 ? String(data[r][iMemo]) : '';
    if (fav || memo.trim()) out.push({ name: String(data[r][iName]).trim(), lat: data[r][iLat], lng: data[r][iLng], fav, memo });
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function noteText_(s) {
  return ContentService.createTextOutput(s);
}

// 開頭是 = + - @ 的字串前面加單引號，避免被當成公式
function noteSanitize_(v) {
  v = String(v || '');
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
