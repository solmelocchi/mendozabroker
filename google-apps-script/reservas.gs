// Reuniones pedidas desde la web: una fila por reunión en la hoja «Reservas».
// La web lee de acá qué horarios ya están tomados y los muestra en gris.
// Borrar una fila libera ese horario.
//
// Cómo se instala (una sola vez):
//   1. Crear una hoja de cálculo nueva en Google Sheets.
//   2. Extensiones › Apps Script, borrar lo que haya y pegar este archivo.
//   3. Implementar › Nueva implementación › tipo «Aplicación web»,
//      «Ejecutar como: yo» y «Quién tiene acceso: cualquier usuario».
//   4. Copiar la URL que termina en /exec y pegarla en RESERVAS_URL de index.html.

const HOJA = "Reservas";

function hoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA);
  if (!h) {
    h = libro.insertSheet(HOJA);
    h.appendRow(["Fecha", "Hora", "Zona", "Nombre", "Apellido", "Ciudad y país", "Teléfono", "Correo", "Tema", "Empresa", "Cargado"]);
  }
  return h;
}

// solo fecha y hora: la web no necesita saber quién pidió cada reunión
function ocupados() {
  return hoja().getDataRange().getDisplayValues().slice(1).map(f => f[0] + " " + f[1]);
}

const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  return json({ ocupados: ocupados() });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const candado = LockService.getScriptLock();
  candado.waitLock(10000);                         // dos pedidos al mismo tiempo no toman el mismo horario
  try {
    if (ocupados().includes(d.fecha + " " + d.hora)) return json({ ok: false, motivo: "ocupado" });
    // el apóstrofo deja la fecha y la hora como texto: si no, Sheets las convierte y dejan de coincidir
    hoja().appendRow(["'" + d.fecha, "'" + d.hora, d.zona, d.nombre, d.apellido, d.ciudad_y_pais, d.telefono, d.email, d.tema, d.empresa, new Date()]);
    return json({ ok: true });
  } finally {
    candado.releaseLock();
  }
}
