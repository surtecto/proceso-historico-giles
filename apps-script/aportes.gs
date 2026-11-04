/**
 * Línea histórica de San Andrés de Giles: receptor de los aportes de las familias.
 *
 * Para ponerlo en marcha (una sola vez):
 * 1. Crear una Hoja de cálculo de Google nueva, por ejemplo «Línea histórica de Giles · aportes».
 * 2. Extensiones > Apps Script. Borrar lo que haya, pegar todo este código y guardar.
 * 3. Implementar > Nueva implementación > tipo «Aplicación web».
 *    Ejecutar como: yo. Quién tiene acceso: cualquier usuario. Autorizar cuando lo pida.
 * 4. Copiar la URL que termina en /exec y ponerla en la web (aportes_config.json → "url").
 *
 * Cómo se usa:
 * - Cada aporte llega como una fila de la hoja «Aportes», con Estado «pendiente».
 *   Las fotos quedan en la carpeta de Drive «Línea histórica de Giles · fotos de aportes», privadas.
 * - Para publicarlo, elegir «publicado» en la columna Estado. Antes se puede corregir el año,
 *   el título o el texto. En unos minutos aparece en la línea, en el círculo de su época,
 *   y sus fotos pasan a verse con el enlace. «descartado» lo deja fuera.
 * - El nombre y el contacto de quien lo manda nunca salen de la hoja.
 */
var HOJA = 'Aportes';
var CARPETA = 'Línea histórica de Giles · fotos de aportes';
var COLS = ['Recibido', 'Estado', 'Id', 'Año', 'Fecha', 'Tipo', 'Título', 'Qué pasó', 'Familia o grupo', 'Lugar',
            'Autor de las fotos', 'Fotos (ids de Drive)', 'Nombre (no se publica)', 'Contacto (no se publica)'];

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA); h.appendRow(COLS); h.setFrozenRows(1);
    h.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['pendiente', 'publicado', 'descartado'], true).build());
  }
  return h;
}
function carpeta_() { var it = DriveApp.getFoldersByName(CARPETA); return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA); }
function txt_(v, n) { return String(v == null ? '' : v).slice(0, n); }
function salida_(s, json) {
  var o = ContentService.createTextOutput(s);
  return json ? o.setMimeType(ContentService.MimeType.JSON) : o;
}

function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var p = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (p.web) return salida_('ok');                       // campo trampa contra robots
    var y = parseInt(p.y, 10);
    if (!(y >= 1700 && y <= 2100) || !p.t) return salida_('incompleto');
    var id = 'a' + Utilities.getUuid().slice(0, 8), ids = [];
    (p.fotos || []).slice(0, 3).forEach(function (f, i) {
      if (!f || !f.datos) return;
      var blob = Utilities.newBlob(Utilities.base64Decode(f.datos), 'image/jpeg', id + '-' + (i + 1) + '.jpg');
      ids.push(carpeta_().createFile(blob).getId());
    });
    hoja_().appendRow([new Date(), 'pendiente', id, y, "'" + txt_(p.fecha, 10), txt_(p.tipo, 40), txt_(p.t, 90), txt_(p.d, 600),
                       txt_(p.familia, 80), txt_(p.lugar, 80), txt_(p.autor, 80), ids.join(' '), txt_(p.nombre, 60), txt_(p.contacto, 80)]);
    CacheService.getScriptCache().remove('publicados');
    return salida_('ok');
  } finally { lock.releaseLock(); }
}

function doGet() {
  var cache = CacheService.getScriptCache(), c = cache.get('publicados');
  if (c) return salida_(c, true);
  var aportes = [];
  hoja_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[1]).trim().toLowerCase() !== 'publicado') return;
    var y = parseInt(r[3], 10); if (!(y >= 1700 && y <= 2100)) return;
    var fecha = r[4] instanceof Date ? Utilities.formatDate(r[4], 'GMT-3', 'yyyy-MM-dd') : String(r[4] || y);
    if (fecha.slice(0, 4) !== String(y)) fecha = String(y);   // si se corrigió el año, manda el año
    var fotos = String(r[11] || '').split(/\s+/).filter(String).map(function (fid) {
      try {
        var f = DriveApp.getFileById(fid);
        if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return {url: 'https://drive.google.com/thumbnail?id=' + fid + '&sz=w1600'};
      } catch (err) { return null; }
    }).filter(Boolean);
    aportes.push({id: String(r[2]), y: y, fecha: fecha, tipo: String(r[5]), t: String(r[6]), d: String(r[7]),
                  familia: String(r[8]), lugar: String(r[9]), autor: String(r[10]), fotos: fotos, nfotos: fotos.length});
  });
  var s = JSON.stringify({aportes: aportes});
  cache.put('publicados', s, 300);
  return salida_(s, true);
}
