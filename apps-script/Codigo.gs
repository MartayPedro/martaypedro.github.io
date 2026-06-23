/**
 * Backend de la web de boda Marta & Pedro José.
 *
 * Conecta el formulario de confirmación (RSVP) y las sugerencias de música
 * con una hoja de cálculo de Google Sheets (dos pestañas: "RSVP" y "Playlist").
 * Además sirve los datos al panel de control privado (protegido con clave).
 *
 * CÓMO INSTALARLO / ACTUALIZARLO:
 *  1. Crea (o abre) la hoja de cálculo en Google Sheets.
 *  2. Menú: Extensiones → Apps Script.
 *  3. Borra lo que haya y pega TODO este archivo. Guarda.
 *  4. CAMBIA la clave del panel abajo (CLAVE_PANEL) por una tuya.
 *  5. Despliega: Implementar → Gestionar implementaciones → (editar la
 *     existente) → Nueva versión, o "Nueva implementación" → "Aplicación web".
 *       - Ejecutar como: Yo mismo.
 *       - Quién tiene acceso: Cualquier persona.
 *  6. La URL /exec ya está puesta en la web. Si creas una nueva, pásasela a Claude.
 *
 * Las pestañas y cabeceras se crean automáticamente la primera vez.
 */

var HOJA_RSVP = "RSVP";
var HOJA_PLAYLIST = "Playlist";

// ⚠️ CAMBIA esta clave por una tuya. Es la contraseña para entrar al panel.
var CLAVE_PANEL = "marta-pedro-2026";

function doGet(e) {
  try {
    // Si llega la clave correcta, devolvemos TODOS los datos para el panel.
    if (e && e.parameter && e.parameter.panel) {
      if (e.parameter.panel !== CLAVE_PANEL) {
        return json({ ok: false, error: "Clave incorrecta" });
      }
      return json({
        ok: true,
        rsvp: leerRSVP(),
        canciones: leerCanciones(),
        resumen: calcularResumen()
      });
    }
    // Petición pública: solo la playlist.
    return json({ ok: true, canciones: leerCanciones() });
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

/* ---------- LECTURA ---------- */

function leerCanciones() {
  var hoja = obtenerHoja(HOJA_PLAYLIST);
  var valores = hoja.getDataRange().getValues();
  var canciones = [];
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    if (!fila[0]) continue;
    canciones.push({
      id: String(fila[0]),
      titulo: fila[1],
      artista: fila[2],
      proponente: fila[3],
      votos: Number(fila[4]) || 0
    });
  }
  return canciones;
}

function leerRSVP() {
  var hoja = obtenerHoja(HOJA_RSVP);
  var valores = hoja.getDataRange().getValues();
  var filas = [];
  for (var i = 1; i < valores.length; i++) {
    var f = valores[i];
    if (!f[1]) continue; // sin nombre, fila vacía
    filas.push({
      fecha: f[0] ? Utilities.formatDate(new Date(f[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "",
      nombre: f[1],
      asiste: f[2],
      personas: f[3],
      acompanantes: f[4],
      ninos: f[5],
      bebes: f[6],
      alergias: f[7],
      comentarios: f[8]
    });
  }
  return filas;
}

function calcularResumen() {
  var rsvp = leerRSVP();
  var r = {
    respuestas: rsvp.length,
    asistenSi: 0,
    asistenNo: 0,
    totalComensales: 0,
    totalNinos: 0,
    totalBebes: 0
  };
  rsvp.forEach(function (x) {
    if (String(x.asiste) === "si") {
      r.asistenSi++;
      r.totalComensales += Number(x.personas) || 0;
      r.totalNinos += Number(x.ninos) || 0;
      r.totalBebes += Number(x.bebes) || 0;
    } else if (String(x.asiste) === "no") {
      r.asistenNo++;
    }
  });
  return r;
}

/* ---------- ESCRITURA ---------- */

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

/* ---------- UTILIDADES ---------- */

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
