/**
 * Lalu: terima pendaftaran beta dari landing page dan simpan ke Google Sheet.
 * Projek Apps Script berdiri sendiri; Sheet dibuka ikut ID.
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Salin URL Web app ke ENDPOINT dalam index.html.
 */

var SHEET_ID = '1Y7F2n3QoA4Oj-ZCQD3Ecl3M7kz02FVBCCLEzHlmGqBw';
var HEADERS = ['Masa', 'Kod', 'Nama', 'WhatsApp', 'E-mel', 'Negeri', 'Kenderaan',
               'Jenis', 'Syarikat', 'Armada', 'Sumber (utm)', 'Kempen (utm)', 'Dirujuk oleh', 'Setuju dihubungi'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'sibuk' });
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Perangkap bot: medan tersembunyi mesti kosong
    if (d.website) return json_({ ok: true });

    if (d.type === 'tempat') return addPlace_(d);

    var nama = clean_(d.nama, 80);
    var wa = clean_(d.whatsapp, 20).replace(/[^\d+]/g, '');
    var emel = clean_(d.emel, 120);
    if (!nama || (!wa && !emel) || d.setuju !== true) return json_({ ok: false, error: 'tidak lengkap' });

    var sheet = getSheet_();

    // Elak pendaftaran berganda dengan nombor atau e-mel yang sama
    var last = sheet.getLastRow();
    if (last > 1) {
      var rows = sheet.getRange(2, 4, last - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        if ((wa && String(rows[i][0]).replace(/^'/, '') === wa) || (emel && String(rows[i][1]).toLowerCase() === emel.toLowerCase())) {
          return json_({ ok: true, duplicate: true });
        }
      }
    }

    sheet.appendRow([
      new Date(),
      safe_(clean_(d.kod, 10)),
      safe_(nama),
      wa ? "'" + wa : '',
      safe_(emel),
      safe_(clean_(d.negeri, 40)),
      safe_(clean_(d.kenderaan, 20)),
      d.jenis === 'Syarikat' ? 'Syarikat' : 'Peribadi',
      safe_(clean_(d.syarikat, 120)),
      safe_(clean_(d.armada, 20)),
      safe_(clean_(d.sumber, 60)),
      safe_(clean_(d.kempen, 60)),
      safe_(clean_(d.rujukan, 10)),
      'Ya'
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'ralat' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Tempat: cadangan pengguna, diluluskan pemilik ---------- */
var PLACE_SHEET = 'Tempat';
var PLACE_HEADERS = ['Masa', 'ID', 'Nama tempat', 'Kategori', 'Lat', 'Lng', 'Nota', 'Status'];
var PLACE_CATS = ['Pejabat / bisnes', 'Kedai / restoran', 'Masjid / surau', 'Sekolah / institusi', 'Taman / perumahan', 'Lain-lain'];

function getPlaceSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(PLACE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PLACE_SHEET);
    sheet.getRange(1, 1, 1, PLACE_HEADERS.length).setValues([PLACE_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Menunggu', 'Lulus', 'Tolak'], true).build();
    sheet.getRange(2, 8, 1000, 1).setDataValidation(rule);
  }
  return sheet;
}

function addPlace_(d) {
  var nama = clean_(d.nama, 80);
  var lat = Number(d.lat), lng = Number(d.lng);
  if (!nama || !isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    return json_({ ok: false, error: 'tidak lengkap' });
  }
  var cat = PLACE_CATS.indexOf(d.kategori) >= 0 ? d.kategori : 'Lain-lain';
  var sheet = getPlaceSheet_();
  sheet.appendRow([
    new Date(),
    Utilities.getUuid().slice(0, 8),
    safe_(nama),
    cat,
    Math.round(lat * 1e6) / 1e6,
    Math.round(lng * 1e6) / 1e6,
    safe_(clean_(d.nota, 200)),
    'Menunggu'
  ]);
  return json_({ ok: true });
}

// GET ?type=tempat: senarai tempat berstatus Lulus untuk carian app
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.type !== 'tempat') return json_({ ok: true });
  var cache = CacheService.getScriptCache();
  var hit = cache.get('tempat');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var sheet = getPlaceSheet_();
  var last = sheet.getLastRow();
  var out = [];
  if (last > 1) {
    var rows = sheet.getRange(2, 1, last - 1, PLACE_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[7]).trim().toLowerCase() !== 'lulus') continue;
      var lat = Number(r[4]), lng = Number(r[5]);
      if (!r[2] || !isFinite(lat) || !isFinite(lng)) continue;
      out.push({ id: String(r[1]), nama: String(r[2]).replace(/^'/, ''), kategori: String(r[3]), lat: lat, lng: lng, nota: String(r[6]).replace(/^'/, '') });
    }
  }
  var body = JSON.stringify({ ok: true, tempat: out });
  if (body.length < 90000) cache.put('tempat', body, 300);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'Masa') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  if (sheet.getFrozenRows() !== 1) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

// Jalankan dari editor: beri kebenaran, tetapkan zon waktu Malaysia, buang baris ujian
function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ss.setSpreadsheetTimeZone('Asia/Kuala_Lumpur');
  var sheet = getSheet_();
  for (var r = sheet.getLastRow(); r >= 2; r--) {
    if (String(sheet.getRange(r, 3).getValue()).indexOf('UJIAN Claude') === 0) sheet.deleteRow(r);
  }
  var places = getPlaceSheet_();
  for (var q = places.getLastRow(); q >= 2; q--) {
    if (String(places.getRange(q, 3).getValue()).indexOf('UJIAN Claude') === 0) places.deleteRow(q);
  }
  CacheService.getScriptCache().remove('tempat');
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max);
}

// Elak formula injection: nilai yang bermula dengan = + - @ tidak dibaca sebagai formula
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
