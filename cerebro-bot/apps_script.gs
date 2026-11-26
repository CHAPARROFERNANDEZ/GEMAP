/**
 * SEGUNDO CEREBRO - puente entre Google Drive y el bot de Railway.
 * Se ejecuta en la cuenta de Google de Yuri (script.google.com).
 *
 * Qué hace cada minuto:
 *  1. vigilarDrive(): mira si hay archivos nuevos o modificados en 1_ENTRADA de alguna empresa
 *     y se lo comunica al bot (POST /drive-event).
 *  2. enviarAvisos(): lee la hoja "Avisos" del TABLERO CENTRAL y manda al bot (POST /aviso)
 *     las filas que aún no se han enviado.
 *
 * Propiedades del script (Configuración del proyecto > Propiedades del script):
 *   BOT_URL        https://<dominio del servicio en Railway>
 *   SHARED_SECRET  el mismo valor que la variable SHARED_SECRET de Railway
 * Servicios avanzados: añadir "Drive API" (identificador Drive, versión v3).
 */
var EMPRESAS_ID = '1J7BjVqtD1LzypOHlAXYUr5PEXvj_L2Az';
var TABLERO_ID = '1QNkjsfn3X52nwB6mXXQVQmx4V-wiB0AD0zi_u0V0MUM';

function instalar() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('vigilarDrive').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('enviarAvisos').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('moverArchivos').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('publicarPendientes').timeBased().everyMinutes(5).create();
}

function post_(path, payload) {
  var p = PropertiesService.getScriptProperties();
  var r = UrlFetchApp.fetch(p.getProperty('BOT_URL') + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Secret': p.getProperty('SHARED_SECRET') },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return r.getResponseCode() === 200;
}

// Devuelve {empresa, entrada} si la carpeta cuelga de EMPRESAS/<empresa>/..., o null.
function ubicar_(folderId) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('f_' + folderId);
  if (hit) return hit === 'null' ? null : JSON.parse(hit);
  var path = [], cur = folderId, depth = 0;
  while (cur && cur !== EMPRESAS_ID && depth < 8) {
    var f = Drive.Files.get(cur, { fields: 'id,name,parents' });
    path.unshift(f.name);
    cur = f.parents && f.parents.length ? f.parents[0] : null;
    depth++;
  }
  var res = (cur === EMPRESAS_ID && path.length >= 2)
    ? { empresa: path[0], entrada: path[1] === '1_ENTRADA' } : null;
  cache.put('f_' + folderId, res ? JSON.stringify(res) : 'null', 6 * 3600);
  return res;
}

function vigilarDrive() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('PAGE_TOKEN');
  if (!token) {
    props.setProperty('PAGE_TOKEN', Drive.Changes.getStartPageToken().startPageToken);
    return;
  }
  var porEmpresa = {}, newToken = token, page = token;
  while (page) {
    var r = Drive.Changes.list(page, {
      pageSize: 100,
      fields: 'nextPageToken,newStartPageToken,changes(fileId,removed,file(id,name,parents,mimeType,trashed,modifiedTime))'
    });
    (r.changes || []).forEach(function (c) {
      var f = c.file;
      if (c.removed || !f || f.trashed || f.mimeType === 'application/vnd.google-apps.folder') return;
      if (!f.parents || !f.parents.length) return;
      var u = ubicar_(f.parents[0]);
      if (!u || !u.entrada) return;
      (porEmpresa[u.empresa] = porEmpresa[u.empresa] || []).push(
        { name: f.name, id: f.id, modified: f.modifiedTime });
    });
    if (r.nextPageToken) { page = r.nextPageToken; } else { newToken = r.newStartPageToken || token; page = null; }
  }
  var ok = true;
  Object.keys(porEmpresa).forEach(function (emp) {
    if (!post_('/drive-event', { empresa: emp, archivos: porEmpresa[emp] })) ok = false;
  });
  // Solo avanza el marcador si el bot recibió todo; si falla, se reintenta el minuto siguiente.
  if (ok) props.setProperty('PAGE_TOKEN', newToken);
}

function enviarAvisos() {
  var sh = SpreadsheetApp.openById(TABLERO_ID).getSheetByName('Avisos');
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r[3] && !r[4]) {
      if (post_('/aviso', { empresa: r[1], prioridad: r[2], texto: r[3] })) {
        sh.getRange(i + 2, 5).setValue(new Date());
      }
    }
  }
}


// El agente pide mover archivos escribiendo filas en la hoja "Movimientos"
// (Fecha, Empresa, Id archivo, Nombre archivo, Id carpeta destino, Ruta destino, Motivo, Hecho).
// Este script, que se ejecuta con tu cuenta, los mueve. Solo dentro de la misma empresa.
function moverArchivos() {
  var sh = SpreadsheetApp.openById(TABLERO_ID).getSheetByName('Movimientos');
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[2] || !r[4] || r[7]) continue;
    var estado;
    try {
      var f = Drive.Files.get(String(r[2]).trim(), { fields: 'id,parents,trashed' });
      var origen = f.parents && f.parents.length ? ubicar_(f.parents[0]) : null;
      var destino = ubicar_(String(r[4]).trim());
      if (f.trashed) {
        estado = 'RECHAZADO: archivo en la papelera';
      } else if (!origen || !destino || origen.empresa !== destino.empresa) {
        estado = 'RECHAZADO: origen y destino no son de la misma empresa';
      } else if (f.parents[0] === String(r[4]).trim()) {
        estado = 'OK (ya estaba en su sitio)';
      } else {
        DriveApp.getFileById(f.id).moveTo(DriveApp.getFolderById(String(r[4]).trim()));
        estado = 'OK ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm');
      }
    } catch (e) {
      estado = 'ERROR: ' + String(e.message).slice(0, 120);
    }
    sh.getRange(i + 2, 8).setValue(estado);
  }
}

// Manda al bot las tareas abiertas agrupadas por empresa, para consultarlas por Telegram.
function publicarPendientes() {
  var sh = SpreadsheetApp.openById(TABLERO_ID).getSheetByName('Tareas');
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var por = {};
  rows.forEach(function (r) {
    if (!r[0] || !r[1]) return;
    if (String(r[5]).toLowerCase().indexOf('cerr') === 0) return;
    var lim = r[4] instanceof Date ? Utilities.formatDate(r[4], 'Europe/Madrid', 'dd/MM/yyyy') : String(r[4] || '');
    (por[r[1]] = por[r[1]] || []).push({ id: r[0], tarea: String(r[2]).slice(0, 900), prio: String(r[3]), limite: lim });
  });
  post_('/pendientes', { empresas: por });
  publicarAsientos_();
}

// Publica la hoja Asientos para que el bot genere el Excel (/asientos).
function publicarAsientos_() {
  var sh = SpreadsheetApp.openById(TABLERO_ID).getSheetByName('Asientos');
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  var por = {};
  rows.forEach(function (r) {
    if (!r[0] || !r[1]) return;
    var f = r[2] instanceof Date ? Utilities.formatDate(r[2], 'Europe/Madrid', 'dd/MM/yyyy') : String(r[2] || '');
    var num = function (v) { return v === '' || v === null ? '' : Number(v); };
    (por[r[0]] = por[r[0]] || []).push([String(r[1]), f, String(r[4]), num(r[5]), num(r[6]), String(r[7]), String(r[8]), String(r[3])]);
  });
  post_('/asientos', { empresas: por });
}
