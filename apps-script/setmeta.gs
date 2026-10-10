/* ============================================================
   reel2map 網頁寫回「收藏」「我的說明」用的 action=setmeta
   安裝方式：
   1. Apps Script 編輯器左邊「檔案 +」→ 指令碼，命名 setmeta，貼上這整個檔案
   2. 在原本 Code.gs 的 doGet(e) 第一行（token 檢查之前）加上：
        if (e.parameter.action === 'setmeta') return setMeta_(e);
   3. 管理部署作業 → 既有部署的鉛筆 → 版本選「新版本」→ 部署
      （不要按「新增部署作業」，會換網址）

   NOTE_KEY 會寫在公開的 index.html 裡，所以它只能改這兩欄，
   不能新增/刪除地點；原本的 SECRET 不需要也不應該放進網頁。
   ============================================================ */

const NOTE_KEY = 'r2m-note-7Kq3vX9p';

function setMeta_(e) {
  const p = e.parameter;
  if (p.key !== NOTE_KEY && p.token !== SECRET) return text_('forbidden');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(DATA_SPREADSHEET_ID).getSheets()[0];
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
    if (row < 0) return text_('not found');

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
      cell.setValue(sanitize_(p.memo));
    }
    return text_('ok');
  } finally {
    lock.releaseLock();
  }
}

function text_(s) {
  return ContentService.createTextOutput(s);
}
