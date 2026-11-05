// Reuniones pedidas desde la web: una fila por reunión, en la pestaña del país.
// Chile va a MULTIRISK y Argentina a WORLDLOGISTIC: cada una tiene su agenda, así
// que el mismo horario se puede pedir una vez en cada país.
// La web lee de acá qué horarios ya están tomados y los muestra en gris.
// Borrar una fila libera ese horario.
//
// Cómo se instala (una sola vez):
//   1. Crear una hoja de cálculo nueva en Google Sheets.
//   2. Extensiones › Apps Script, borrar lo que haya y pegar este archivo.
//   3. Implementar › Nueva implementación › tipo «Aplicación web»,
//      «Ejecutar como: yo» y «Quién tiene acceso: cualquier usuario».
//   4. Copiar la URL que termina en /exec y pegarla en RESERVAS_URL de index.html.
// Para cambiar el código después: Implementar › Gestionar implementaciones, el
// lápiz, «Nueva versión». Así la dirección sigue siendo la misma.

const PESTANAS = { cl: "MULTIRISK", ar: "WORLDLOGISTIC" };
const TITULOS = ["Fecha", "Hora", "Zona", "Nombre", "Apellido", "Ciudad y país", "Teléfono", "Correo", "Tema", "Empresa", "Cargado"];

// la pestaña del país, sin importar mayúsculas ni espacios en el nombre
function hoja(pais) {
  const nombre = PESTANAS[pais];
  if (!nombre) throw new Error("País desconocido: " + pais);
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const igual = n => n.replace(/\s+/g, "").toUpperCase() === nombre;
  const h = libro.getSheets().find(s => igual(s.getName())) || libro.insertSheet(nombre);
  if (h.getRange(1, 1).getDisplayValue() !== "Fecha") {
    if (h.getLastRow() > 0) h.insertRowBefore(1);
    h.getRange(1, 1, 1, TITULOS.length).setValues([TITULOS]).setFontWeight("bold");
  }
  return h;
}

// solo fecha y hora: la web no necesita saber quién pidió cada reunión
function ocupados(pais) {
  return hoja(pais).getDataRange().getDisplayValues().slice(1).filter(f => f[0]).map(f => f[0] + " " + f[1]);
}

const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  return json({ ocupados: ocupados(e.parameter.pais || "cl") });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents), pais = d.pais || "cl";
  const candado = LockService.getScriptLock();
  candado.waitLock(10000);                         // dos pedidos al mismo tiempo no toman el mismo horario
  try {
    if (ocupados(pais).includes(d.fecha + " " + d.hora)) return json({ ok: false, motivo: "ocupado" });
    // el apóstrofo deja la fecha y la hora como texto: si no, Sheets las convierte y dejan de coincidir
    hoja(pais).appendRow(["'" + d.fecha, "'" + d.hora, d.zona, d.nombre, d.apellido, d.ciudad_y_pais, d.telefono, d.email, d.tema, d.empresa, new Date()]);
    return json({ ok: true });
  } finally {
    candado.releaseLock();
  }
}
