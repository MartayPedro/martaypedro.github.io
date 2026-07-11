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
var HOJA_CONFIG = "Config"; // guarda la distribución de mesas (JSON en A1)

// ⚠️ CAMBIA esta clave por una tuya. Es la contraseña para entrar al panel.
var CLAVE_PANEL = "marta-pedro-2026";

// ⚠️ Clave de la API de Google Gemini para el asistente "Azahar".
//   - Consíguela GRATIS en https://aistudio.google.com/apikey
//   - Pégala aquí ENTRE LAS COMILLAS (se queda en tu Apps Script, nunca en la web pública).
//   - Si la dejas vacía, el asistente avisa de que no está configurado.
var GEMINI_API_KEY = "";
var GEMINI_MODEL = "gemini-2.5-flash";

// Personalidad e información que usa el asistente Azahar.
var PROMPT_AZAHAR =
  'Eres Azahar, la distinguida y simpática coordinadora de bodas de Marta y Pedro José. ' +
  'Te diriges a los invitados con inmensa cercanía, cariño y elegancia, empleando giros sutiles del acento andaluz de Córdoba sin perder la pulcritud y el tono romántico minimalista del evento. ' +
  'Tienes pleno conocimiento de la planificación de la boda para el 11 de Octubre de 2026.\n\n' +
  'DATOS CRUCIALES DE LA BODA:\n' +
  '- Pareja: Marta y Pedro José.\n' +
  '- Fecha del enlace: Domingo 11 de Octubre de 2026.\n' +
  '- Ceremonia: A las 17:30 H en la "Iglesia de Santa Marina de Aguas Santas" en Córdoba capital (Plaza de Santa Marina, s/n).\n' +
  '- Banquete / Fiesta: A partir de las 19:30 H en la "Hacienda S\'cultura" (Carretera Palma del Río, Km 6.8, 14005 Córdoba). Se recomienda coger taxi, autobús habilitado o coche particular.\n' +
  '- Alojamiento concertado: "Hotel Puerta Osario" (Calle Osario, 7). Código promocional "TRENADO" con descuento; se aplica en la web oficial del hotel al seleccionar fechas. Estancia mínima 2 noches. Teléfono: +34 957 485 411.\n' +
  '- Fecha límite para confirmar la asistencia: 1 de Septiembre de 2026 en el formulario de esta web.\n' +
  '- En esta web los invitados también pueden proponer música bailable y subir fotos.\n\n' +
  'HISTORIA DE AMOR: Se conocieron entre amigos y miradas; una noche de septiembre bajo el cielo de Córdoba empezaron su historia; el "sí, quiero" fue en las Maldivas, frente al océano.\n\n' +
  'Responde siempre de manera cálida, sucinta pero muy clara. Si te preguntan temas ajenos a la boda, recuérdalo con cariño y redirige al gran día de Marta y Pedro José.';

function doGet(e) {
  try {
    // Asistente Azahar: resolvemos el chat en el servidor (la clave nunca sale de aquí).
    if (e && e.parameter && e.parameter.chat) {
      return json(responderAzahar(e.parameter.chat, e.parameter.historia));
    }
    // Si llega la clave correcta, devolvemos TODOS los datos para el panel.
    if (e && e.parameter && e.parameter.panel) {
      if (e.parameter.panel !== CLAVE_PANEL) {
        return json({ ok: false, error: "Clave incorrecta" });
      }
      return json({
        ok: true,
        rsvp: leerRSVP(),
        canciones: leerCanciones(),
        resumen: calcularResumen(),
        mesas: leerMesas()
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
    } else if (data.tipo === "editar") {
      // Edición protegida con la clave del panel.
      if (data.clave !== CLAVE_PANEL) return json({ ok: false, error: "Clave incorrecta" });
      editarCelda(data.fila, data.campo, data.valor);
    } else if (data.tipo === "borrarFamilia") {
      if (data.clave !== CLAVE_PANEL) return json({ ok: false, error: "Clave incorrecta" });
      if (data.fila) obtenerHoja(HOJA_RSVP).deleteRow(Number(data.fila));
    } else if (data.tipo === "guardarMesas") {
      if (data.clave !== CLAVE_PANEL) return json({ ok: false, error: "Clave incorrecta" });
      obtenerHoja(HOJA_CONFIG).getRange(1, 1).setValue(JSON.stringify(data.mesas || {}));
    } else if (data.tipo === "editarCancion") {
      if (data.clave !== CLAVE_PANEL) return json({ ok: false, error: "Clave incorrecta" });
      editarCancion(data.id, data.campo, data.valor);
    } else if (data.tipo === "borrarCancion") {
      if (data.clave !== CLAVE_PANEL) return json({ ok: false, error: "Clave incorrecta" });
      borrarCancion(data.id);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Edita una celda concreta de la pestaña RSVP.
function editarCelda(fila, campo, valor) {
  var columnas = {
    nombre: 2, asiste: 3, personas: 4, acompanantes: 5,
    ninos: 6, bebes: 7, alergias: 8, comentarios: 9,
    ninosNombres: 10, bebesNombres: 11, notas: 12
  };
  var col = columnas[campo];
  if (!col || !fila) return;
  obtenerHoja(HOJA_RSVP).getRange(Number(fila), col).setValue(valor);
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
      fila: i + 1,
      fecha: f[0] ? Utilities.formatDate(new Date(f[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "",
      nombre: f[1],
      asiste: f[2],
      personas: f[3],
      acompanantes: f[4],
      ninos: f[5],
      bebes: f[6],
      alergias: f[7],
      comentarios: f[8],
      ninosNombres: f[9],
      bebesNombres: f[10],
      notas: f[11]
    });
  }
  return filas;
}

// Llama a Gemini con el mensaje y el historial reciente. Devuelve {ok, reply} o {ok:false, error}.
function responderAzahar(mensaje, historiaJson) {
  if (!GEMINI_API_KEY) return { ok: false, error: "sin-clave" };
  var contents = [];
  try { if (historiaJson) contents = JSON.parse(historiaJson); } catch (e) { contents = []; }
  if (!(contents instanceof Array)) contents = [];
  contents.push({ role: "user", parts: [{ text: String(mensaje || "") }] });

  var payload = {
    contents: contents,
    systemInstruction: { parts: [{ text: PROMPT_AZAHAR }] }
  };
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL + ":generateContent?key=" + GEMINI_API_KEY;
  var resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    return { ok: false, error: "IA " + resp.getResponseCode() };
  }
  var data = JSON.parse(resp.getContentText());
  var texto = data && data.candidates && data.candidates[0] && data.candidates[0].content &&
    data.candidates[0].content.parts && data.candidates[0].content.parts[0].text;
  return { ok: true, reply: texto || "" };
}

function leerMesas() {
  var hoja = obtenerHoja(HOJA_CONFIG);
  var valor = hoja.getRange(1, 1).getValue();
  if (!valor) return { tables: [], asignaciones: {} };
  try {
    var obj = JSON.parse(valor);
    if (!obj.tables) obj.tables = [];
    if (!obj.asignaciones) obj.asignaciones = {};
    return obj;
  } catch (e) {
    return { tables: [], asignaciones: {} };
  }
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
      "Niños (menú infantil)", "Bebés (tronas)", "Alergias por comensal", "Comentarios",
      "Nombres niños", "Nombres bebés", "Notas (novios)"
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
    data.comentarios || "",
    data.ninosNombres || "",
    data.bebesNombres || "",
    ""
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

// Edita un campo de una canción (búsqueda por id).
function editarCancion(id, campo, valor) {
  var columnas = { titulo: 2, artista: 3, proponente: 4, votos: 5 };
  var col = columnas[campo];
  if (!col) return;
  var hoja = obtenerHoja(HOJA_PLAYLIST);
  var valores = hoja.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      hoja.getRange(i + 1, col).setValue(campo === "votos" ? (Number(valor) || 0) : valor);
      return;
    }
  }
}

// Borra una canción (búsqueda por id).
function borrarCancion(id) {
  var hoja = obtenerHoja(HOJA_PLAYLIST);
  var valores = hoja.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      hoja.deleteRow(i + 1);
      return;
    }
  }
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
