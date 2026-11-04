// Sigga Einars ferðir – bókunarþjónn
// Límdu þetta í Apps Script sem er bundið við skjalið "Sigga Einars ferðir – Bókanir"
// (Extensions → Apps Script), vistaðu, og Deploy → New deployment → Web app:
//   Execute as: Me   ·   Who has access: Anyone
// Afritaðu Web app URL í SCRIPT_URL í index.html.

var SHEET_NAME = 'Sheet1';
var FROM_NAME = 'Sigga Einars ferðir';
var BUS_SEATS = 11;

var STOPS = {
  'Hvolsvöllur': '14:00 (N1)',
  'Selfoss': '14:50 (Olís)',
  'Hveragerði': '15:20 (N1)',
  'Reykjavík': '16:20 (Olís Norðlingaholti)',
  'Olís Norðlingaholti': '16:20 (Olís Norðlingaholti)'
};

function doPost(e) {
  var p = (e && e.parameter) || {};
  var nafn = String(p.nafn || 'Farþegi').trim().slice(0, 80);
  var netfang = String(p.netfang || '').trim().slice(0, 120);
  var stopp = String(p.stopp || 'Hvolsvöllur').trim().slice(0, 60);
  var saeti = String(p.saeti || 'Miðja').trim().slice(0, 60);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var n = Math.max(sh.getLastRow(), 1); // header is row 1
  var midi = 'SEF-' + ('000' + n).slice(-3);
  var stada = n > BUS_SEATS ? 'Kerran' : 'Staðfest';

  sh.appendRow([new Date(), midi, nafn, netfang, stopp, saeti, stada]);

  if (netfang && /@/.test(netfang)) {
    try {
      MailApp.sendEmail({
        to: netfang,
        name: FROM_NAME,
        subject: 'Farmiði ' + midi + ' – Húsafell, það skýrist',
        htmlBody: ticketHtml(midi, nafn, stopp, saeti, stada),
        body: 'Farmiði ' + midi + ' fyrir ' + nafn + '. Brottför ' + stopp + ' kl. ' + (STOPS[stopp] || 'það skýrist') + '. Sæti: ' + saeti + '. Staða: ' + stada + '.'
      });
    } catch (err) {
      sh.getRange(sh.getLastRow(), 7).setValue(stada + ' (póstur mistókst)');
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, midi: midi, stada: stada }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('Sigga Einars ferðir – bókunarþjónn er í gangi. Það skýrist.');
}

function ticketHtml(midi, nafn, stopp, saeti, stada) {
  var timi = STOPS[stopp] || 'það skýrist';
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  return '' +
  '<div style="background:#f8f4ea;padding:32px 16px;font-family:Nunito,Helvetica,Arial,sans-serif;color:#1b3a36">' +
  '  <div style="max-width:560px;margin:0 auto;background:#ffffff;border:2px solid #e0a63a;border-radius:18px;overflow:hidden">' +
  '    <div style="background:#1b3a36;color:#ffffff;padding:28px 28px 24px">' +
  '      <div style="display:inline-block;background:#e0a63a;color:#1b3a36;font-size:12px;font-weight:900;letter-spacing:1.5px;text-transform:uppercase;padding:6px 14px;border-radius:999px">Sigga Einars ferðir</div>' +
  '      <div style="font-size:30px;font-weight:900;margin-top:14px;line-height:1.1">Farmiði</div>' +
  '      <div style="color:#cfe3dc;font-weight:700;margin-top:6px">Humlar og húðfellingar í Húsafelli</div>' +
  '    </div>' +
  '    <table style="width:100%;border-collapse:collapse;font-size:16px">' +
  row('Miðanúmer', '<strong style="font-size:22px;letter-spacing:1px">' + esc(midi) + '</strong>') +
  row('Farþegi', esc(nafn)) +
  row('Brottför', esc(stopp) + ' · kl. ' + esc(timi)) +
  row('Sæti', esc(saeti)) +
  row('Koma í Húsafell', 'Það skýrist') +
  row('Staða', '<strong>' + esc(stada) + '</strong>') +
  '    </table>' +
  '    <div style="padding:20px 28px;border-top:2px dashed #dfe6e2;font-size:13px;color:#5f746f;font-weight:600">' +
  '      Miðinn er ekki endurgreiddur. Söngbókin er eitt lag. Sá sem gleymist á Olís fær viðurkenningarskjal.' +
  '    </div>' +
  '    <div style="background:#1b3a36;color:#ffffff;text-align:center;padding:16px;font-weight:900;letter-spacing:1.5px;text-transform:uppercase;font-size:14px">Húsafell – það skýrist!</div>' +
  '  </div>' +
  '</div>';

  function row(k, v) {
    return '<tr><td style="padding:12px 28px;border-bottom:1px solid #eef3ef;color:#5f746f;font-size:12px;font-weight:800;letter-spacing:1.2px;text-transform:uppercase;width:38%">' + k + '</td>' +
           '<td style="padding:12px 28px 12px 0;border-bottom:1px solid #eef3ef;font-weight:700">' + v + '</td></tr>';
  }
}
