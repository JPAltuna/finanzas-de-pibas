/**
 * FINANZAS DE PIBAS — Conector Apps Script
 * ─────────────────────────────────────────
 * Este script conecta la app "Finanzas de pibas" con TU planilla de Google.
 * Cada usuaria pega este código en el Apps Script de SU planilla y lo
 * implementa como Web App. La app le habla por la URL que termina en /exec.
 *
 * CÓMO INSTALARLO (una sola vez):
 *  1. Creá una planilla de Google nueva (nombre sugerido: "Finanzas de pibas").
 *  2. Menú Extensiones → Apps Script.
 *  3. Borrá lo que haya y pegá TODO este archivo. Guardá (ícono disquete).
 *  4. Botón azul "Implementar" → "Nueva implementación".
 *     - Tipo: Aplicación web
 *     - Ejecutar como: Yo
 *     - Quién tiene acceso: Cualquier usuario
 *  5. Autorizá con tu cuenta de Google (va a pedir permisos, es tu propio script).
 *  6. Copiá la URL que termina en /exec y pegala en la app.
 *
 * Las pestañas (GASTOS, INGRESOS, INVERSIONES, CONFIG) se crean solas
 * la primera vez que la app se conecta. Podés editar las categorías y
 * cuentas directamente en la pestaña CONFIG.
 */

var HOJAS = {
  GASTOS: ['ID', 'FECHA', 'MONTO', 'CATEGORIA', 'TIPO', 'MEDIO', 'NOTA', 'REGISTRADO'],
  INGRESOS: ['ID', 'FECHA', 'MONTO', 'CATEGORIA', 'CUENTA', 'NOTA', 'REGISTRADO'],
  INVERSIONES: ['ID', 'FECHA', 'TIPO', 'PLATAFORMA', 'ACTIVO', 'MONTO_USD', 'CANTIDAD', 'NOTA', 'REGISTRADO']
};

var CONFIG_SEED = {
  'CATEGORIAS GASTOS': ['Comida', 'Súper', 'Casa', 'Transporte', 'Salud', 'Ropa', 'Salidas', 'Regalos', 'Suscripciones', 'Educación', 'Mascotas', 'Otros'],
  'CATEGORIAS INGRESOS': ['Sueldo', 'Freelance', 'Ventas', 'Regalo', 'Reintegro', 'Otros'],
  'CUENTAS': ['Efectivo', 'Banco', 'Mercado Pago', 'Ualá', 'Brubank', 'Dólares en mano']
};

// ───────────────────────── GET (lecturas, JSONP) ─────────────────────────
function doGet(e) {
  var p = (e && e.parameter) || {};
  var accion = p.accion || 'ping';
  var out;
  try {
    if (accion === 'ping') out = { ok: true, app: 'finanzas-de-pibas', version: 1 };
    else if (accion === 'config') out = { ok: true, config: leerConfig() };
    else if (accion === 'datos') out = { ok: true, datos: leerDatos() };
    else if (accion === 'confirmar') out = { ok: true, existe: existeId(String(p.id || ''), String(p.hoja || '')) };
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
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var accion = body.accion || '';
    if (accion === 'gasto') out = agregarFila('GASTOS', body, ['fecha', 'monto', 'categoria', 'tipo', 'medio', 'nota']);
    else if (accion === 'ingreso') out = agregarFila('INGRESOS', body, ['fecha', 'monto', 'categoria', 'cuenta', 'nota']);
    else if (accion === 'inversion') out = agregarFila('INVERSIONES', body, ['fecha', 'tipo', 'plataforma', 'activo', 'monto_usd', 'cantidad', 'nota']);
    else if (accion === 'config_nueva') out = configAgregar(body.lista, body.valor);
    else if (accion === 'config_borrar') out = configBorrar(body.lista, body.valor);
    else out = { ok: false, error: 'Acción desconocida: ' + accion };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (ignorar) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── Núcleo ─────────────────────────
function hoja(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    var cab = HOJAS[nombre];
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else {
    asegurarColumnas(h, HOJAS[nombre]);
  }
  return h;
}

// Si esta versión del conector suma una columna nueva (ej. TIPO en GASTOS),
// la inserta en su lugar en las hojas ya creadas, sin tocar los datos.
function asegurarColumnas(h, cab) {
  var actual = h.getRange(1, 1, 1, Math.max(h.getLastColumn(), 1)).getValues()[0]
    .map(function (v) { return String(v).trim(); });
  for (var i = 0; i < cab.length; i++) {
    if (actual[i] === cab[i]) continue;
    if (actual.indexOf(cab[i]) >= 0) continue; // está en otra posición: no tocar
    h.insertColumns(i + 1, 1);
    h.getRange(1, i + 1).setValue(cab[i]).setFontWeight('bold');
    actual.splice(i, 0, cab[i]);
  }
}

function hojaConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName('CONFIG');
  if (!h) {
    h = ss.insertSheet('CONFIG');
    var cols = Object.keys(CONFIG_SEED);
    h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    for (var c = 0; c < cols.length; c++) {
      var vals = CONFIG_SEED[cols[c]].map(function (v) { return [v]; });
      h.getRange(2, c + 1, vals.length, 1).setValues(vals);
    }
    h.setFrozenRows(1);
  }
  return h;
}

function leerConfig() {
  var h = hojaConfig();
  var datos = h.getDataRange().getValues();
  var cab = datos[0];
  var out = {};
  for (var c = 0; c < cab.length; c++) {
    var lista = [];
    for (var f = 1; f < datos.length; f++) {
      var v = String(datos[f][c] || '').trim();
      if (v) lista.push(v);
    }
    out[String(cab[c])] = lista;
  }
  return {
    categoriasGastos: out['CATEGORIAS GASTOS'] || [],
    categoriasIngresos: out['CATEGORIAS INGRESOS'] || [],
    cuentas: out['CUENTAS'] || []
  };
}

function leerDatos() {
  hojaConfig(); // asegura que CONFIG exista desde la primera conexión
  var out = {};
  Object.keys(HOJAS).forEach(function (nombre) {
    var h = hoja(nombre);
    var cab = HOJAS[nombre];
    var n = h.getLastRow();
    var filas = [];
    if (n > 1) {
      var vals = h.getRange(2, 1, n - 1, cab.length).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (!String(vals[i][0])) continue; // sin ID = fila vacía
        var obj = {};
        for (var c = 0; c < cab.length; c++) {
          var v = vals[i][c];
          if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          obj[cab[c].toLowerCase()] = v;
        }
        filas.push(obj);
      }
    }
    out[nombre.toLowerCase()] = filas;
  });
  return out;
}

function agregarFila(nombre, body, campos) {
  var h = hoja(nombre);
  var id = String(body.id || '').trim();
  if (!id) return { ok: false, error: 'Falta el id del registro' };
  if (existeId(id, nombre)) return { ok: true, id: id, repetido: true }; // ya estaba (reintento), no duplicar

  var fila = [id];
  for (var i = 0; i < campos.length; i++) {
    var v = body[campos[i]];
    fila.push(v === undefined || v === null ? '' : v);
  }
  fila.push(new Date());
  h.getRange(h.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
  return { ok: true, id: id };
}

// Agrega una categoría o cuenta nueva a la pestaña CONFIG (desde la app).
// Para sacar opciones que no usás, borrá la celda directamente en CONFIG.
var CONFIG_COLUMNAS = {
  categoriasGastos: 'CATEGORIAS GASTOS',
  categoriasIngresos: 'CATEGORIAS INGRESOS',
  cuentas: 'CUENTAS'
};
function configAgregar(lista, valor) {
  var nombreCol = CONFIG_COLUMNAS[String(lista || '')];
  valor = String(valor || '').trim();
  if (!nombreCol) return { ok: false, error: 'Lista desconocida: ' + lista };
  if (!valor) return { ok: false, error: 'Falta el valor' };
  var h = hojaConfig();
  var datos = h.getDataRange().getValues();
  var col = -1;
  for (var c = 0; c < datos[0].length; c++) {
    if (String(datos[0][c]).trim() === nombreCol) { col = c; break; }
  }
  if (col < 0) return { ok: false, error: 'No encontré la columna ' + nombreCol + ' en CONFIG' };
  var ultima = 1;
  for (var f = 1; f < datos.length; f++) {
    var v = String(datos[f][col] || '').trim();
    if (v.toLowerCase() === valor.toLowerCase()) return { ok: true, valor: v, repetido: true };
    if (v) ultima = f + 1;
  }
  h.getRange(ultima + 1, col + 1).setValue(valor);
  return { ok: true, valor: valor };
}

// Saca una categoría o cuenta de la pestaña CONFIG (desde la app).
// Los registros viejos que la usaban no se tocan: solo desaparece del desplegable.
function configBorrar(lista, valor) {
  var nombreCol = CONFIG_COLUMNAS[String(lista || '')];
  valor = String(valor || '').trim();
  if (!nombreCol) return { ok: false, error: 'Lista desconocida: ' + lista };
  if (!valor) return { ok: false, error: 'Falta el valor' };
  var h = hojaConfig();
  var datos = h.getDataRange().getValues();
  var col = -1;
  for (var c = 0; c < datos[0].length; c++) {
    if (String(datos[0][c]).trim() === nombreCol) { col = c; break; }
  }
  if (col < 0) return { ok: false, error: 'No encontré la columna ' + nombreCol + ' en CONFIG' };
  for (var f = 1; f < datos.length; f++) {
    if (String(datos[f][col] || '').trim().toLowerCase() === valor.toLowerCase()) {
      h.getRange(f + 1, col + 1).deleteCells(SpreadsheetApp.Dimension.ROWS);
      return { ok: true, valor: valor };
    }
  }
  return { ok: true, valor: valor, noEstaba: true };
}

function existeId(id, nombre) {
  if (!id || !HOJAS[nombre]) return false;
  var h = hoja(nombre);
  var n = h.getLastRow();
  if (n < 2) return false;
  var ids = h.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return true;
  }
  return false;
}
