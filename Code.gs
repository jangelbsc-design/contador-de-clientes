const SPREADSHEET_ID = '1CG6jiQEjqU4FePm94Y2wPSRs6GaI5UIVuI5H4AkUNX0';
const SHEET_NAME = 'Visitas Taller';

// Configurar encabezados si no existen
function setupSheet() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    const newSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_NAME);
    newSheet.appendRow(['ID Turno', 'Fecha', 'Hora Entrada', 'Motivo', 'Estado', 'Hora Atención', 'Tiempo Espera (min)', 'Operador']);
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID Turno', 'Fecha', 'Hora Entrada', 'Motivo', 'Estado', 'Hora Atención', 'Tiempo Espera (min)', 'Operador']);
    }
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'create') {
      return createTicket(data.motivo);
    } else if (action === 'attend') {
      return attendTicket(data.idTurno, data.operador);
    } else if (action === 'finish') {
      return finishTicket(data.idTurno, data.operador);
    }

    return responseJson({ error: 'Acción no válida' });
  } catch (error) {
    return responseJson({ error: error.toString() });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getPending') {
    return getPendingTickets();
  } else if (action === 'getStats') {
    return getStats();
  } else if (action === 'setup') {
    setupSheet();
    return responseJson({ success: true, message: 'Hoja configurada correctamente.' });
  }

  return responseJson({ message: 'DISMAC API en línea. Usa POST para crear/atender/finish, o GET con action=getPending / getStats' });
}

function createTicket(motivo) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  
  const dateObj = new Date();
  const dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  const data = sheet.getDataRange().getValues();
  let countToday = 0;
  
  // Contar cuántos tickets hay hoy para generar el consecutivo
  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][1];
    if (rowDate instanceof Date) {
       rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (rowDate === dateStr) {
      countToday++;
    }
  }
  
  const idTurno = 'T-' + (countToday + 1);
  
  sheet.appendRow([
    idTurno,
    dateStr,
    timeStr,
    motivo,
    'Pendiente',
    '',
    '',
    ''
  ]);
  
  return responseJson({ success: true, idTurno: idTurno, message: '¡Registrado! Te atendemos en breve' });
}

function attendTicket(idTurno, operador) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const dateObj = new Date();
  const timeStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idTurno && data[i][4] === 'Pendiente') {
      let rowDate = data[i][1];
      let rowTime = data[i][2];
      
      let dateStr = rowDate instanceof Date ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : rowDate;
      let timeStrEntrada = rowTime instanceof Date ? Utilities.formatDate(rowTime, Session.getScriptTimeZone(), 'HH:mm:ss') : rowTime;
      
      const dateTimeEntrada = new Date(dateStr + 'T' + timeStrEntrada);
      let waitMinutes = Math.floor((dateObj - dateTimeEntrada) / 60000);
      if (isNaN(waitMinutes)) waitMinutes = 0;
      
      sheet.getRange(i + 1, 5).setValue('En Proceso');
      sheet.getRange(i + 1, 6).setValue(timeStr);
      sheet.getRange(i + 1, 7).setValue(waitMinutes);
      sheet.getRange(i + 1, 8).setValue(operador);
      
      found = true;
      break;
    }
  }
  
  if (found) {
    return responseJson({ success: true, message: 'Turno en proceso por ' + operador });
  } else {
    return responseJson({ success: false, message: 'Turno no encontrado o ya fue atendido' });
  }
}

function finishTicket(idTurno, operador) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    // Validamos que sea el mismo ID, que esté En Proceso y que sea el mismo operador
    if (data[i][0] === idTurno && data[i][4] === 'En Proceso' && data[i][7] === operador) {
      sheet.getRange(i + 1, 5).setValue('Finalizado');
      found = true;
      break;
    }
  }
  
  if (found) {
    return responseJson({ success: true, message: 'Turno finalizado' });
  } else {
    return responseJson({ success: false, message: 'Turno no encontrado o no autorizado' });
  }
}

function getPendingTickets() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return responseJson({ tickets: [] });

  const data = sheet.getDataRange().getValues();
  const pending = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === 'Pendiente' || data[i][4] === 'En Proceso') {
      let rowDate = data[i][1];
      let rowTime = data[i][2];
      
      let dateStr = rowDate instanceof Date ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : rowDate;
      let timeStr = rowTime instanceof Date ? Utilities.formatDate(rowTime, Session.getScriptTimeZone(), 'HH:mm:ss') : rowTime;
      
      let entryTime = 0;
      try {
        entryTime = new Date(dateStr + 'T' + timeStr).getTime();
      } catch (e) {
        entryTime = Date.now(); 
      }

      pending.push({
        idTurno: data[i][0],
        motivo: data[i][3],
        estado: data[i][4],
        operador: data[i][7], // Importante para saber quién lo tiene En Proceso
        entryTimestamp: entryTime
      });
    }
  }
  
  return responseJson({ tickets: pending });
}

function getStats() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return responseJson({ totalVisitas: 0, promedioEspera: 0 });

  const data = sheet.getDataRange().getValues();
  const dateStrToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  let count = 0;
  let totalWait = 0;
  let waitCount = 0;

  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][1];
    if (rowDate instanceof Date) {
       rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    if (rowDate === dateStrToday) {
      count++;
      let wait = data[i][6];
      if (wait !== '' && wait !== null && !isNaN(wait)) {
        totalWait += Number(wait);
        waitCount++;
      }
    }
  }
  
  const avg = waitCount > 0 ? Math.round(totalWait / waitCount) : 0;
  
  return responseJson({ totalVisitas: count, promedioEspera: avg });
}

function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Función para enviar el balance diario por correo
// IMPORTANTE: Configurar un Activador (Trigger) en Apps Script para que se ejecute a una hora específica (ej. 20:00)
function sendDailyBalanceEmail() {
  const emailDestino = Session.getActiveUser().getEmail(); // O pon tu correo aquí entre comillas: "tucorreo@email.com"
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const dateStrToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  let totalAtendidos = 0;
  let operatorCounts = {};

  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][1];
    if (rowDate instanceof Date) {
       rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    // Contar solo las visitas de hoy que hayan sido atendidas o finalizadas
    if (rowDate === dateStrToday && (data[i][4] === 'En Proceso' || data[i][4] === 'Finalizado' || data[i][4] === 'Atendido')) {
      totalAtendidos++;
      let op = data[i][7]; // Columna Operador
      if (op) {
        if (!operatorCounts[op]) operatorCounts[op] = 0;
        operatorCounts[op]++;
      }
    }
  }

  // Si no hubo nadie, no enviamos nada o enviamos un reporte vacío
  if (totalAtendidos === 0) {
    MailApp.sendEmail({
      to: emailDestino,
      subject: 'Reporte Diario DISMAC - ' + dateStrToday,
      htmlBody: '<h3>Balance Diario DISMAC</h3><p>Hoy no se registraron clientes atendidos.</p>'
    });
    return;
  }

  // Encontrar al operador que más atendió
  let bestOperator = '';
  let maxAtenciones = 0;
  
  for (let op in operatorCounts) {
    if (operatorCounts[op] > maxAtenciones) {
      maxAtenciones = operatorCounts[op];
      bestOperator = op;
    }
  }

  // Construir el cuerpo del correo
  let body = '<h2>Balance Diario del Taller DISMAC - ' + dateStrToday + '</h2>';
  body += '<p><strong>Total de clientes atendidos:</strong> ' + totalAtendidos + '</p>';
  body += '<h3>Resumen por Dispatcher / Operador:</h3><ul>';
  
  for (let op in operatorCounts) {
    body += '<li>' + op + ': ' + operatorCounts[op] + ' tickets</li>';
  }
  body += '</ul>';
  
  body += '<br><p>🏆 <strong>El operador que más atendió hoy fue:</strong> ' + bestOperator + ' (con ' + maxAtenciones + ' atenciones).</p>';

  // Enviar el correo
  MailApp.sendEmail({
    to: emailDestino,
    subject: 'Reporte Diario DISMAC - ' + dateStrToday,
    htmlBody: body
  });
}
