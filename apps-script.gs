const SHEET_NAME = 'Respuestas';

const ENSAYOS = [
  ['2026-07-13-nico', 'Lunes 13 de julio — Uno a uno con Nico'],
  ['2026-07-15-grupal-nico', 'Miércoles 15 de julio — Ensayo grupal con Nico, Jefe, compañero de trabajo, amigo de Jona y Negro'],
  ['2026-07-16-lorena', 'Jueves 16 de julio — Uno a uno con Lorena'],
  ['2026-07-17-llamada', 'Viernes 17 de julio — Escena de la llamada — Nico + Jona'],
  ['2026-07-20-jona', 'Lunes 20 de julio — Uno a uno con Jona'],
  ['2026-07-21-nico-nino', 'Martes 21 de julio — Uno a uno con Nico niño'],
  ['2026-07-22-jona-nino', 'Miércoles 22 de julio — Uno a uno con Jona niño'],
  ['2026-07-27-abuela', 'Lunes 27 de julio — Uno a uno con Abuela'],
  ['2026-07-28-ninos', 'Martes 28 de julio — Ensayo Nico niño + Jona niño'],
  ['2026-07-29-protas', 'Miércoles 29 de julio — Ensayo de protagonistas — Nico, Jona, Lorena y Abuela'],
  ['2026-07-31-locacion', 'Viernes 31 de julio, noche en locación — Ensayo Nico + Jona']
];

function doPost(e) {
  const data = JSON.parse(e.postData.contents || '{}');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headers = ['Timestamp', 'Nombre', 'Rol', 'Contacto', 'Comentarios'].concat(ENSAYOS.map(item => item[1]));

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  const row = [new Date(), data.nombre || '', data.rol || '', data.contacto || '', data.comentarios || ''];
  ENSAYOS.forEach(function(item) {
    const id = item[0];
    row.push((data.respuestas && data.respuestas[id] && data.respuestas[id].respuesta) || '');
  });

  sheet.appendRow(row);
  sheet.autoResizeColumns(1, headers.length);
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
