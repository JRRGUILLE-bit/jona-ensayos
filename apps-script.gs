const RAW_SHEET_NAME = 'Respuestas_RAW';
const SUMMARY_SHEET_NAME = 'Resumen';

const PERSONAJES = [
  { actor: 'Dante', personaje: 'Alexis' },
  { actor: 'Martín', personaje: 'Negro' },
  { actor: 'Maxi', personaje: 'Amigo de Jona' },
  { actor: 'Lucas', personaje: 'Compañero de taller de Nico' },
  { actor: 'Caro', personaje: 'Lorena' },
  { actor: 'Guille', personaje: 'Jona' },
  { actor: 'Tere', personaje: 'Abuela' },
  { actor: 'Nico', personaje: 'Nico' },
  { actor: 'Camilo', personaje: 'Nico niño' },
  { actor: 'Aún no definido', personaje: 'Jona niño' }
];

const ENSAYOS = [
  { id: '2026-07-13-nico', fecha: 'Lunes 13 de julio', actividad: 'Uno a uno con Nico', para: ['Nico'] },
  { id: '2026-07-15-grupal-nico', fecha: 'Miércoles 15 de julio', actividad: 'Ensayo Nico + Alexis + Compañero de taller de Nico + Amigo de Jona + Negro', para: ['Nico', 'Alexis', 'Compañero de taller de Nico', 'Amigo de Jona', 'Negro'] },
  { id: '2026-07-16-lorena', fecha: 'Jueves 16 de julio', actividad: 'Uno a uno con Lorena', para: ['Lorena'] },
  { id: '2026-07-17-llamada', fecha: 'Viernes 17 de julio', actividad: 'Escena de la llamada — Nico + Jona', para: ['Nico', 'Jona'] },
  { id: '2026-07-20-jona', fecha: 'Lunes 20 de julio', actividad: 'Uno a uno con Jona', para: ['Jona'] },
  { id: '2026-07-21-nico-nino', fecha: 'Martes 21 de julio', actividad: 'Uno a uno con Nico niño', para: ['Nico niño'] },
  { id: '2026-07-22-jona-nino', fecha: 'Miércoles 22 de julio', actividad: 'Uno a uno con Jona niño', para: ['Jona niño'] },
  { id: '2026-07-27-abuela', fecha: 'Lunes 27 de julio', actividad: 'Uno a uno con Abuela', para: ['Abuela'] },
  { id: '2026-07-28-ninos', fecha: 'Martes 28 de julio', actividad: 'Ensayo Nico niño + Jona niño', para: ['Nico niño', 'Jona niño'] },
  { id: '2026-07-29-protas', fecha: 'Miércoles 29 de julio', actividad: 'Ensayo de protagonistas — Nico, Jona, Lorena y Abuela', para: ['Nico', 'Jona', 'Lorena', 'Abuela'] },
  { id: '2026-07-31-locacion', fecha: 'Viernes 31 de julio, noche en locación', actividad: 'Ensayo Nico + Jona', para: ['Nico', 'Jona'] }
];

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!e || !e.postData || !e.postData.contents) {
    ensureRawSheet_(ss);
    rebuildSummary_(ss);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, mode: 'setup' })).setMimeType(ContentService.MimeType.JSON);
  }

  const data = JSON.parse(e.postData.contents || '{}');
  appendRawResponse_(ss, data);
  rebuildSummary_(ss);

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRawSheet_(ss);
  rebuildSummary_(ss);
  return ContentService.createTextOutput('Resumen actualizado desde Respuestas_RAW').setMimeType(ContentService.MimeType.TEXT);
}

function setupNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRawSheet_(ss);
  rebuildSummary_(ss);
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === RAW_SHEET_NAME && e.range.getRow() > 1) {
    rebuildSummary_(SpreadsheetApp.getActiveSpreadsheet());
  }
}

function appendRawResponse_(ss, data) {
  const sheet = ensureRawSheet_(ss);
  const row = [new Date(), data.personaje || '', data.actor || actorFor_(data.personaje), data.comentarios || ''];

  ENSAYOS.forEach(function(ensayo) {
    const respuesta = data.respuestas && data.respuestas[ensayo.id] && data.respuestas[ensayo.id].respuesta ? data.respuestas[ensayo.id].respuesta : '';
    row.push(respuesta);
  });

  sheet.appendRow(row);
  sheet.autoResizeColumns(1, row.length);
}

function ensureRawSheet_(ss) {
  let sheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(RAW_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    const headers = ['Timestamp', 'Personaje', 'Actor/a', 'Comentarios'].concat(ENSAYOS.map(function(ensayo) {
      return ensayo.fecha + ' — ' + ensayo.actividad;
    }));

    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#24584d').setFontColor('#ffffff');
  }

  return sheet;
}

function rebuildSummary_(ss) {
  const raw = ensureRawSheet_(ss);
  let summary = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summary) summary = ss.insertSheet(SUMMARY_SHEET_NAME);

  summary.clear();
  summary.getRange(1, 1, summary.getMaxRows(), summary.getMaxColumns()).breakApart();

  const rawValues = raw.getDataRange().getValues();
  const latestByPersonaje = {};

  for (let i = 1; i < rawValues.length; i++) {
    const row = rawValues[i];
    const personaje = row[1];
    if (personaje) latestByPersonaje[personaje] = row;
  }

  const actorHeaders = PERSONAJES.map(function(p) {
    return p.personaje + ' — ' + p.actor;
  });

  const table = [];
  table.push(['Fecha', 'Ensayo'].concat(actorHeaders).concat(['Pueden', 'No pueden', 'Pendientes']));

  ENSAYOS.forEach(function(ensayo, ensayoIndex) {
    let pueden = 0;
    let noPueden = 0;
    let pendientes = 0;
    const row = [ensayo.fecha, ensayo.actividad];

    PERSONAJES.forEach(function(p) {
      if (ensayo.para.indexOf(p.personaje) === -1) {
        row.push('—');
        return;
      }

      const rawRow = latestByPersonaje[p.personaje];
      const value = rawRow ? rawRow[4 + ensayoIndex] : '';

      if (value === 'Puedo') {
        pueden++;
        row.push('Puedo');
      } else if (value === 'No puedo') {
        noPueden++;
        row.push('No puedo');
      } else {
        pendientes++;
        row.push('Pendiente');
      }
    });

    row.push(pueden, noPueden, pendientes);
    table.push(row);
  });

  const totalColumns = table[0].length;

  summary.getRange(1, 1).setValue('Resumen disponibilidad ensayos — Jona').setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#24584d');
  summary.getRange(1, 1, 1, totalColumns).setBackground('#24584d');
  summary.getRange(2, 1).setValue('Actualizado desde Respuestas_RAW: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')).setFontStyle('italic').setFontColor('#607068');

  summary.getRange(4, 1, table.length, totalColumns).setValues(table);

  const backgrounds = table.map(function(row, rowIndex) {
    return row.map(function(cell) {
      if (rowIndex === 0) return '#24584d';
      if (cell === 'Puedo') return '#d9ead3';
      if (cell === 'No puedo') return '#f4cccc';
      if (cell === 'Pendiente') return '#fff2cc';
      if (cell === '—') return '#eeeeee';
      return '#fffdf8';
    });
  });

  summary.getRange(4, 1, table.length, totalColumns).setBackgrounds(backgrounds).setWrap(true).setVerticalAlignment('middle');
  summary.getRange(4, 1, 1, totalColumns).setFontWeight('bold').setFontColor('#ffffff');
  summary.getRange(5, 3, table.length - 1, PERSONAJES.length).setHorizontalAlignment('center');
  summary.getRange(5, totalColumns - 2, table.length - 1, 3).setHorizontalAlignment('center').setFontWeight('bold');

  summary.setFrozenRows(4);
  summary.setFrozenColumns(2);
  summary.setColumnWidth(1, 150);
  summary.setColumnWidth(2, 360);

  for (let c = 3; c <= totalColumns; c++) {
    summary.setColumnWidth(c, 130);
  }

  ss.setActiveSheet(summary);
  ss.moveActiveSheet(1);

  const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (rawSheet) {
    ss.setActiveSheet(rawSheet);
    ss.moveActiveSheet(2);
  }

  ss.setActiveSheet(summary);
}

function actorFor_(personaje) {
  const found = PERSONAJES.find(function(p) {
    return p.personaje === personaje;
  });

  return found ? found.actor : '';
}
