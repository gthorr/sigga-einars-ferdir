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

// ---------------------------------------------------------------------------
// Tilkynningar til allra bókaðra farþega – sendar sem "Sigga Einars ferðir".
// Keyrðu sendFarangurslisti() úr ritlinum (Run ▸ sendFarangurslisti). Ekki þarf
// að deploya aftur; fyrsta skiptið biður Google um leyfi fyrir póstsendingu.
// ---------------------------------------------------------------------------

var SKIP_NAMES = ['Maggi mix', 'Prufa Prufudottir']; // fá ekki tilkynningar

function sendFarangurslisti() {
  var html = noticeHtml('Farangurslisti', 'Eyðublað nr. 3 · Leið nr. 1 · Húsafell',
    '<p style="margin:0 0 14px;font-weight:700">Kæri farþegi,</p>' +
    '<p style="margin:0 0 18px">Farangurslisti fyrir Leið nr. 1 – Humlar og húðfellingar í Húsafelli – hefur verið birtur. Farþegi skal hafa meðferðis:</p>' +
    checkRow('Sængurföt', 'Eins og venjulega.') +
    checkRow('Sundskýla', 'Eins og venjulega.') +
    checkRow('Annað', 'Segir sig sjálft.', true) +
    '<p style="margin:18px 0 0;font-size:13px;color:#5f746f;font-weight:600">§1 Farangur fer í kerruna. Farþegar fara ekki í kerruna. Spyrjið Ævar Örn ef eitthvað er óljóst hér! ' +
    '§2 Söngbók nr. 1 verður ekki afhent í rútunni. Gott að taka hana með að heiman. ' +
    '§3 Farþegi sem gleymir sængurfötum sefur á sundskýlunni. ' +
    '§4 Farþegi sem gleymir sundskýlu getur haft samband við Denna, hann gæti verið með auka.</p>',
    'https://gthorr.github.io/sigga-einars-ferdir/#farangur', 'Opna farangurslistann');
  var text = 'Kæri farþegi,\n\nFarangurslisti fyrir Leið nr. 1 hefur verið birtur. Farþegi skal hafa meðferðis:\n[ ] Sængurföt (eins og venjulega)\n[ ] Sundskýla (eins og venjulega)\nAnnað: segir sig sjálft.\n\nhttps://gthorr.github.io/sigga-einars-ferdir/#farangur\n\nFarangur fer í kerruna. Farþegar fara ekki í kerruna.\n\nSigga Einars ferðir\nRúta: já · Kerra: já · Endurgreiðsla: nei';
  return sendNotice('Sigga Einars ferðir – Eyðublað nr. 3: Farangurslisti', html, text);
}

// Almenn tilkynning: sendNotice(efni, html, texti) → fjöldi sendra pósta
function sendNotice(subject, html, text) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var rows = sh.getDataRange().getValues();
  var seen = {}, sent = 0;
  for (var i = 1; i < rows.length; i++) {
    var nafn = String(rows[i][2] || '').trim();
    var netfang = String(rows[i][3] || '').trim().toLowerCase();
    if (!netfang || !/@/.test(netfang) || seen[netfang]) continue;
    if (SKIP_NAMES.some(function (s) { return s.toLowerCase() === nafn.toLowerCase(); })) continue;
    seen[netfang] = true;
    MailApp.sendEmail({ to: netfang, name: FROM_NAME, subject: subject, htmlBody: html, body: text });
    sent++;
  }
  Logger.log('Sent ' + sent + ' notices');
  return sent;
}

function checkRow(title, note, dashed) {
  var b = dashed ? '2px dashed #b7c4bf' : '2px solid #1b3a36';
  var c = dashed ? '#5f746f' : '#1b3a36';
  return '<table style="width:100%;border-collapse:separate;border-spacing:0;margin:0 0 10px;border:' + b + ';border-radius:12px;color:' + c + '"><tr>' +
    '<td style="width:36px;padding:14px 0 14px 16px"><div style="width:26px;height:26px;border:' + b + ';border-radius:7px"></div></td>' +
    '<td style="padding:12px 16px"><div style="font-size:20px;font-weight:900">' + title + '</div><div style="font-size:13px;font-weight:600;color:#5f746f">' + note + '</div></td></tr></table>';
}

function noticeHtml(title, sub, bodyHtml, url, btn) {
  return '' +
  '<div style="background:#f8f4ea;padding:32px 16px;font-family:Nunito,Helvetica,Arial,sans-serif;color:#1b3a36">' +
  '  <div style="max-width:560px;margin:0 auto;background:#ffffff;border:2px solid #e0a63a;border-radius:18px;overflow:hidden">' +
  '    <div style="background:#1b3a36;color:#ffffff;padding:28px 28px 24px">' +
  '      <div style="display:inline-block;background:#e0a63a;color:#1b3a36;font-size:12px;font-weight:900;letter-spacing:1.5px;text-transform:uppercase;padding:6px 14px;border-radius:999px">Sigga Einars ferðir</div>' +
  '      <div style="font-size:30px;font-weight:900;margin-top:14px;line-height:1.1">' + title + '</div>' +
  '      <div style="color:#cfe3dc;font-weight:700;margin-top:6px">' + sub + '</div>' +
  '    </div>' +
  '    <div style="padding:24px 28px;font-size:16px;line-height:1.5">' + bodyHtml +
  (url ? '<p style="margin:22px 0 0"><a href="' + url + '" style="display:inline-block;background:#e0a63a;color:#1b3a36;font-weight:900;letter-spacing:1px;text-transform:uppercase;font-size:13px;padding:12px 22px;border-radius:999px;text-decoration:none">' + btn + '</a></p>' : '') +
  '    </div>' +
  '    <div style="padding:16px 28px;border-top:2px dashed #dfe6e2;font-size:13px;color:#5f746f;font-weight:600">Rúta: já · Kerra: já · Endurgreiðsla: nei</div>' +
  '    <div style="background:#1b3a36;color:#ffffff;text-align:center;padding:16px;font-weight:900;letter-spacing:1.5px;text-transform:uppercase;font-size:14px">Húsafell – það skýrist!</div>' +
  '  </div>' +
  '</div>';
}
