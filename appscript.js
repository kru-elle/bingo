function doGet(e) {
  const p = e.parameter;
  const action = p.action || '';

  try {
    if (action === 'save')          return handleSave(p);
    if (action === 'get_players')   return handleGetPlayers();
    if (action === 'get_rankings')  return handleGetRankings();
    return json({ status: 'ok' });
  } catch (err) {
    return json({ error: err.message });
  }
}

const HEADERS = ['วันที่', 'ชื่อ', 'ชั้น', 'เลขที่', 'ข้อที่ใช้', 'ตอบถูก', 'สถานะ'];

function ensureHeader(sh) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#667EEA')
      .setFontColor('white');
  }
}

function handleSave(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bingo');
  ensureHeader(sh);
  const now = new Date().toLocaleString('th-TH');
  sh.appendRow([now, p.name, p.cls, p.num, p.questions, p.correct, p.status]);
  return json({ status: 'saved' });
}

function handleGetPlayers() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bingo');
  ensureHeader(sh);
  const data = sh.getDataRange().getValues();
  const names = [...new Set(data.slice(1).map(r => r[1]).filter(Boolean))];
  return json({ players: names });
}

function handleGetRankings() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bingo');
  ensureHeader(sh);
  const data = sh.getDataRange().getValues();
  return json({ data: data.slice(1) });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}