/**
 * Suniylash order form -> Google Sheet.
 *
 * Setup (once):
 * 1. Open the "Suniylash — Buyurtmalar" sheet -> Extensions -> Apps Script.
 * 2. Replace everything in Code.gs with this file and press Save.
 * 3. Deploy -> New deployment -> type "Web app".
 *    Execute as: Me. Who has access: Anyone. -> Deploy -> allow access.
 * 4. Copy the Web app URL into CONFIG.orderEndpoint in index.html.
 */
const FIELDS = ["name", "phone", "business", "type", "package", "vlang", "details", "lang"];

function doPost(e) {
  const p = (e && e.parameter) || {};
  // Hidden "website" field is only filled in by spam bots.
  if (p.website) return reply({ ok: true });

  // Text starting with = + - @ would be read as a formula (e.g. "+998 ..."), so keep it as text.
  const clean = (v) => {
    const s = String(v || "").slice(0, 1000).trim();
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };
  const row = [new Date()].concat(FIELDS.map((f) => clean(p[f])));
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow(row);

  // Email the sheet owner about the new order.
  try {
    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: "Yangi buyurtma: " + clean(p.name) + " (" + clean(p.business) + ")",
      body: FIELDS.map((f, i) => f + ": " + row[i + 1]).join("\n"),
    });
  } catch (err) {
    // The order is already saved in the sheet; email is a bonus.
  }
  return reply({ ok: true });
}

function reply(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
