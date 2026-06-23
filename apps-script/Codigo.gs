/**
 * Backend de la web de boda Marta & Pedro José.
 *
 * Conecta el formulario de confirmación (RSVP) y las sugerencias de música
 * con una hoja de cálculo de Google Sheets (dos pestañas: "RSVP" y "Playlist").
 *
 * CÓMO INSTALARLO:
 *  1. Crea una hoja de cálculo nueva en Google Sheets.
 *  2. Menú: Extensiones → Apps Script.
 *  3. Borra lo que haya y pega TODO este archivo. Guarda.
 *  4. Despliega: Implementar → Nueva implementación → tipo "Aplicación web".
 *       - Ejecutar como: Yo mismo.
 *       - Quién tiene acceso: Cualquier persona.
 *  5. Copia la URL que termina en /exec y pásasela a Claude (o pégala en
 *     index.html, en la constante APPS_SCRIPT_URL).
 *
 * Las pestañas y cabeceras se crean automáticamente la primera vez.
 */

var HOJA_RSVP = "RSVP";
var HOJA_PLAYLIST = "Playlist";

function doGet(e) {
  // Devuelve las canciones para pintar la playlist en la web.
  try {
    var hoja = obtenerHoja(HOJA_PLAYLIST);
    var valores = hoja.getDataRange().getValues();
    var canciones = [];
    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      if (!fila[0]) continue; // sin id, fila vacía
      canciones.push({
        id: String(fila[0]),
        titulo: fila[1],
        artista: fila[2],
        proponente: fila[3],
        votos: Number(fila[4]) || 0
      });
    }
    return json({ ok: true, canciones: canciones });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.tipo === "rsvp") {
      guardarRSVP(data);
    } else if (data.tipo === "cancion") {
      guardarCancion(data);
    } else if (data.tipo === "voto") {
      sumarVoto(data.id);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function guardarRSVP(data) {
  var hoja = obtenerHoja(HOJA_RSVP);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      "Fecha", "Nombre", "Asiste", "Nº asistentes", "Acompañantes",
      "Niños (menú infantil)", "Bebés (tronas)", "Alergias por comensal", "Comentarios"
    ]);
  }
  hoja.appendRow([
    new Date(),
    data.nombre || "",
    data.asiste || "",
    data.personas || "",
    data.acompanantes || "",
    data.ninos || "0",
    data.bebes || "0",
    formatearAlergias(data.alergias),
    data.comentarios || ""
  ]);
}

function formatearAlergias(alergias) {
  if (!alergias || !alergias.length) return "";
  return alergias.map(function (a) {
    var partes = [];
    if (a.alergias && a.alergias.length) partes.push(a.alergias.join(", "));
    if (a.detalle) partes.push(a.detalle);
    return a.comensal + ": " + partes.join(" / ");
  }).join(" | ");
}

function guardarCancion(data) {
  var hoja = obtenerHoja(HOJA_PLAYLIST);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["id", "Título", "Artista", "Proponente", "Votos"]);
  }
  var id = "c" + new Date().getTime();
  hoja.appendRow([id, data.titulo || "", data.artista || "", data.proponente || "", 1]);
}

function sumarVoto(id) {
  var hoja = obtenerHoja(HOJA_PLAYLIST);
  var valores = hoja.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      var celda = hoja.getRange(i + 1, 5); // columna E (Votos)
      celda.setValue((Number(celda.getValue()) || 0) + 1);
      return;
    }
  }
}

function obtenerHoja(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  return hoja;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
