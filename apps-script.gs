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
