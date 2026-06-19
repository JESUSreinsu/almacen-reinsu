// REINSU Stock — Google Apps Script Backend v2
// Incluye cabeceras CORS para GitHub Pages

const SHEET_NAME = 'Stock';

function doGet(e) {
  const result = handleRequest(e);
  return result;
}

function doPost(e) {
  const result = handleRequest(e);
  return result;
}

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
    
    if (e.parameter && e.parameter.action) {
      action = e.parameter.action;
    }
    
    if (e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
        if (postData.action) action = postData.action;
      } catch(err) {}
    }
    
    let responseData;
    
    if (action === 'getStock') {
      responseData = getStock(sheet);
    } else if (action === 'addMovement' && postData) {
      responseData = addMovement(sheet, postData);
    } else if (action === 'ping') {
      responseData = { ok: true, ts: new Date().toISOString() };
    } else {
      responseData = { error: 'Accion no reconocida: ' + action };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getStock(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { movs: [], stock: {} };
  
  const stock = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    
    const cod   = String(row[1]);
    const tipo  = String(row[3]);
    const qty   = parseFloat(row[4]) || 0;
    
    const mov = {
      ts:    row[0] ? new Date(row[0]).getTime() : Date.now(),
      cod:   cod,
      d:     String(row[2] || ''),
      tipo:  tipo,
      qty:   qty,
      date1: String(row[6] || ''),
      date2: String(row[7] || ''),
      sot:   String(row[8] || ''),
      dst:   String(row[9] || ''),
      notes: String(row[10] || ''),
      user:  String(row[11] || '')
    };
    
    if (!stock[cod]) {
      stock[cod] = { d: mov.d, stock: 0, movs: [] };
    }
    stock[cod].movs.push(mov);
    stock[cod].stock = Math.round(
      (stock[cod].stock + (tipo === 'entrada' ? qty : -qty)) * 1000
    ) / 1000;
  }
  
  return { stock: stock };
}

function addMovement(sheet, data) {
  const mov = data.mov;
  const stockActual = data.stockActual;
  
  if (!mov || !mov.cod) {
    return { error: 'Datos de movimiento incorrectos' };
  }
  
  sheet.appendRow([
    new Date().toISOString(),
    mov.cod,
    mov.d || '',
    mov.tipo,
    mov.qty,
    stockActual,
    mov.date1 || '',
    mov.date2 || '',
    mov.sot || '',
    mov.dst || '',
    mov.notes || '',
    mov.user || ''
  ]);
  
  return { ok: true, stock: stockActual };
}
