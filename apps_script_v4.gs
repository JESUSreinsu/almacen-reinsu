// REINSU Stock — Google Apps Script Backend v4

const SHEET_NAME = 'Stock';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['timestamp','cod','descripcion','tipo','qty','stock_actual','date1','date2','sotano','destino','notas','usuario']);
      sheet.getRange(1,1,1,12).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    }
    let action = '';
    let postData = null;
    if (e.parameter && e.parameter.action) action = e.parameter.action;
    if (e.postData && e.postData.contents) {
      try { postData = JSON.parse(e.postData.contents); if (postData.action) action = postData.action; } catch(err) {}
    }
    let responseData;
    if (action === 'getStock') {
      responseData = getStock(sheet);
    } else if (action === 'addMovement' && postData) {
      responseData = addMovement(sheet, postData);
    } else if (action === 'ping') {
      responseData = { ok: true };
    } else {
      responseData = { error: 'Accion no reconocida: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function formatDate(val) {
  if (!val) return '';
  // If already a string like "2026-06-19", return as is
  if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}/)) return val.slice(0,10);
  // If it's a Date object or date string
  try {
    var d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    var y = d.getFullYear();
    var m = ('0'+(d.getMonth()+1)).slice(-2);
    var day = ('0'+d.getDate()).slice(-2);
    return y+'-'+m+'-'+day;
  } catch(e) { return String(val); }
}

function getStock(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { stock: {} };
  const stock = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    const cod  = String(row[1]).trim();
    const tipo = String(row[3]).trim();
    const qty  = parseFloat(row[4]) || 0;
    if (!stock[cod]) stock[cod] = { d: String(row[2]||''), stock: 0, movs: [] };
    if (tipo === 'entrada') stock[cod].stock += qty;
    else if (tipo === 'consumo') stock[cod].stock -= qty;
    stock[cod].stock = Math.round(stock[cod].stock * 1000) / 1000;
    stock[cod].movs.push({
      ts:    Date.now(),
      tipo:  tipo,
      qty:   qty,
      date1: formatDate(row[6]),
      date2: formatDate(row[7]),
      sot:   String(row[8]||''),
      dst:   String(row[9]||''),
      notes: String(row[10]||''),
      user:  String(row[11]||'')
    });
  }
  return { stock: stock };
}

function addMovement(sheet, data) {
  const mov = data.mov;
  const stockActual = data.stockActual;
  if (!mov || !mov.cod) return { error: 'Datos incorrectos' };
  sheet.appendRow([
    new Date().toISOString(),
    mov.cod, mov.d||'', mov.tipo, mov.qty, stockActual,
    mov.date1||'', mov.date2||'',
    mov.sot||'', mov.dst||'', mov.notes||'', mov.user||''
  ]);
  return { ok: true, stock: stockActual };
}
