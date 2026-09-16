/**
 * FINANZAS DE PIBAS — Conector genérico (versión 2)
 * ──────────────────────────────────────────────────
 * Este script conecta la app "Finanzas de pibas" con TU planilla de Google.
 * Es un conector GENÉRICO: solo sabe leer hojas, agregar filas y editar
 * listas. Toda la lógica vive en la app (que se actualiza sola desde
 * GitHub), así este código casi nunca necesita actualizarse.
 *
 * CÓMO INSTALARLO (una sola vez):
 *  1. Creá una planilla de Google nueva (nombre sugerido: "Finanzas de pibas").
 *  2. Menú Extensiones → Apps Script.
 *  3. Borrá lo que haya y pegá TODO este archivo. Guardá (ícono disquete).
 *  4. Botón azul "Implementar" → "Nueva implementación".
 *     - Tipo: Aplicación web
 *     - Ejecutar como: Yo
 *     - Quién tiene acceso: Cualquier usuario
 *  5. Autorizá con tu cuenta de Google (es tu propio script).
 *  6. Copiá la URL que termina en /exec y pegala en la app.
 *
 * SI YA TENÍAS LA VERSIÓN ANTERIOR: pegá este código encima → guardar →
 * Implementar → Administrar implementaciones → lápiz → Versión: "Nueva
 * versión" → Implementar. La URL /exec no cambia y tus datos quedan igual.
 */

var VERSION = 2;

// ───────────────────────── GET (lecturas, JSONP) ─────────────────────────
function doGet(e) {
  var p = (e && e.parameter) || {};
  var accion = p.accion || 'ping';
  var out;
  try {
    if (accion === 'ping') out = { ok: true, app: 'finanzas-de-pibas', version: VERSION };
    else if (accion === 'leer') out = { ok: true, hojas: leerHojas(String(p.hojas || '')) };
    else if (accion === 'confirmar') out = { ok: true, existe: existeId(String(p.hoja || ''), String(p.id || '')) };
    else out = { ok: false, error: 'Acción desconocida: ' + accion };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  var json = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── POST (escrituras) ─────────────────────────
function doPost(e) {
  var out;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var accion = b.accion || '';
    if (accion === 'fila') out = filaAgregar(b.hoja, b.id, b.esquema, b.datos);
    else if (accion === 'fila_editar') out = filaEditar(b.hoja, b.id, b.datos);
    else if (accion === 'fila_borrar') out = filaBorrar(b.hoja, b.id);
    else if (accion === 'celda_lista_agregar') out = listaAgregar(b.hoja, b.columna, b.valor);
    else if (accion === 'celda_lista_borrar') out = listaBorrar(b.hoja, b.columna, b.valor);
    else if (accion === 'hoja_seed') out = hojaSeed(b.hoja, b.matriz);
    else out = { ok: false, error: 'Acción desconocida: ' + accion };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (ignorar) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── Ayudantes ─────────────────────────
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function cabeceraDe(h) {
  if (h.getLastRow() < 1 || h.getLastColumn() < 1) return [];
  return h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0]
    .map(function (v) { return String(v).trim(); });
}

// Crea la hoja si falta (con el esquema como cabecera) y agrega al final
// las columnas del esquema que no existan. Nunca borra ni mueve nada.
function asegurarHoja(nombre, esquema) {
  nombre = String(nombre || '').trim();
  if (!nombre) throw new Error('Falta el nombre de la hoja');
  var h = ss().getSheetByName(nombre);
  if (!h) {
    h = ss().insertSheet(nombre);
    if (esquema && esquema.length) {
      h.getRange(1, 1, 1, esquema.length).setValues([esquema]).setFontWeight('bold');
      h.setFrozenRows(1);
    }
    return h;
  }
  if (esquema && esquema.length) {
    var cab = cabeceraDe(h);
    for (var i = 0; i < esquema.length; i++) {
      if (cab.indexOf(esquema[i]) < 0) {
        cab.push(esquema[i]);
        h.getRange(1, cab.length).setValue(esquema[i]).setFontWeight('bold');
      }
    }
  }
  return h;
}

function colPorNombre(h, nombre) {
  return cabeceraDe(h).indexOf(String(nombre).trim());
}

function formatearValor(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

// ───────────────────────── Lecturas ─────────────────────────
function leerHojas(listaHojas) {
  var out = {};
  listaHojas.split(',').forEach(function (nombre) {
    nombre = nombre.trim();
    if (!nombre) return;
    var h = ss().getSheetByName(nombre);
    if (!h) { out[nombre] = { cabeceras: [], filas: [] }; return; }
    var cab = cabeceraDe(h);
    var n = h.getLastRow();
    var filas = [];
    if (cab.length && n > 1) {
      var vals = h.getRange(2, 1, n - 1, cab.length).getValues();
      for (var i = 0; i < vals.length; i++) {
        filas.push(vals[i].map(formatearValor));
      }
    }
    out[nombre] = { cabeceras: cab, filas: filas };
  });
  return out;
}

function filaDeId(h, id) {
  var col = colPorNombre(h, 'ID');
  if (col < 0 || !id) return -1;
  var n = h.getLastRow();
  if (n < 2) return -1;
  var ids = h.getRange(2, col + 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function existeId(nombreHoja, id) {
  var h = ss().getSheetByName(nombreHoja);
  if (!h) return false;
  return filaDeId(h, id) > 0;
}

// ───────────────────────── Escrituras ─────────────────────────
function filaAgregar(nombreHoja, id, esquema, datos) {
  id = String(id || '').trim();
  if (!id) return { ok: false, error: 'Falta el id del registro' };
  var h = asegurarHoja(nombreHoja, esquema || []);
  if (filaDeId(h, id) > 0) return { ok: true, id: id, repetido: true }; // reintento: no duplicar
  var cab = cabeceraDe(h);
  var fila = cab.map(function (c) {
    var v = (datos || {})[c];
    return v === undefined || v === null ? '' : v;
  });
  h.getRange(h.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
  return { ok: true, id: id };
}

function filaEditar(nombreHoja, id, datos) {
  var h = ss().getSheetByName(String(nombreHoja || '').trim());
  if (!h) return { ok: false, error: 'No existe la hoja ' + nombreHoja };
  var f = filaDeId(h, id);
  if (f < 0) return { ok: false, error: 'No encontré el registro ' + id };
  var cab = cabeceraDe(h);
  Object.keys(datos || {}).forEach(function (c) {
    var col = cab.indexOf(String(c).trim());
    if (col >= 0) h.getRange(f, col + 1).setValue(datos[c]);
  });
  return { ok: true, id: id };
}

function filaBorrar(nombreHoja, id) {
  var h = ss().getSheetByName(String(nombreHoja || '').trim());
  if (!h) return { ok: false, error: 'No existe la hoja ' + nombreHoja };
  var f = filaDeId(h, id);
  if (f < 0) return { ok: true, id: id, noEstaba: true };
  h.deleteRow(f);
  return { ok: true, id: id };
}

// Listas verticales (ej. las columnas de la hoja CONFIG)
function listaAgregar(nombreHoja, columna, valor) {
  valor = String(valor || '').trim();
  if (!valor) return { ok: false, error: 'Falta el valor' };
  var h = ss().getSheetByName(String(nombreHoja || '').trim());
  if (!h) return { ok: false, error: 'No existe la hoja ' + nombreHoja };
  var col = colPorNombre(h, columna);
  if (col < 0) return { ok: false, error: 'No encontré la columna ' + columna };
  var n = h.getLastRow();
  var ultima = 1;
  if (n > 1) {
    var vals = h.getRange(2, col + 1, n - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0] || '').trim();
      if (v.toLowerCase() === valor.toLowerCase()) return { ok: true, valor: v, repetido: true };
      if (v) ultima = i + 2;
    }
  }
  h.getRange(ultima + 1, col + 1).setValue(valor);
  return { ok: true, valor: valor };
}

function listaBorrar(nombreHoja, columna, valor) {
  valor = String(valor || '').trim();
  if (!valor) return { ok: false, error: 'Falta el valor' };
  var h = ss().getSheetByName(String(nombreHoja || '').trim());
  if (!h) return { ok: false, error: 'No existe la hoja ' + nombreHoja };
  var col = colPorNombre(h, columna);
  if (col < 0) return { ok: false, error: 'No encontré la columna ' + columna };
  var n = h.getLastRow();
  if (n > 1) {
    var vals = h.getRange(2, col + 1, n - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim().toLowerCase() === valor.toLowerCase()) {
        h.getRange(i + 2, col + 1).deleteCells(SpreadsheetApp.Dimension.ROWS);
        return { ok: true, valor: valor };
      }
    }
  }
  return { ok: true, valor: valor, noEstaba: true };
}

// Crea una hoja con contenido inicial (solo si no existe o está vacía)
function hojaSeed(nombreHoja, matriz) {
  nombreHoja = String(nombreHoja || '').trim();
  if (!nombreHoja || !matriz || !matriz.length) return { ok: false, error: 'Faltan datos' };
  var h = ss().getSheetByName(nombreHoja);
  if (h && h.getLastRow() > 0) return { ok: true, yaExiste: true };
  if (!h) h = ss().insertSheet(nombreHoja);
  var ancho = matriz[0].length;
  h.getRange(1, 1, matriz.length, ancho).setValues(matriz);
  h.getRange(1, 1, 1, ancho).setFontWeight('bold');
  h.setFrozenRows(1);
  return { ok: true };
}
