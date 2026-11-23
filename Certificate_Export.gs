/**
 * Certificate Export — LQM Al-Falah 26
 *
 * Bound to the Progress & Certificates spreadsheet (same Apps Script project as
 * Progress_Report.gs — it reuses PROGRESS_TAB, PC_HEADER_ANCHOR and findHeaderRow_()).
 *
 * Writes a CSV of the students who get a certificate, for the certificate-generation
 * script. Source: the Progress_Cert tab. Only rows where column Z = "Y" (case-insensitive)
 * are exported.
 *
 * CSV columns, in this order:
 *   Student ID (E) | Gender (A) | Student Name (H) | Arabic Name (G)
 *
 * Apps Script cannot write to your computer or open a "Save as" window. Instead the menu action
 * opens a small dialog in the browser with a Download button. The browser then saves the file, and
 * asks where when its download setting "Ask where to save each file before downloading" is on
 * (Chrome: Settings → Downloads; Safari asks by default). Nothing is saved to Drive.
 *
 * The file is UTF-8 with a byte-order mark (so Excel shows the Arabic correctly; Python reads it
 * with encoding="utf-8-sig").
 *
 * Read-only for the sheet: nothing in the spreadsheet is changed.
 *
 * Run from the menu: LQM Progress → Export Certificate CSV. (It needs the spreadsheet open in the
 * browser to show the dialog, so running it from the script editor alone will not work.)
 */

// Progress_Cert column numbers (1-indexed)
const CERT_GENDER_COL   = 1;  // A
const CERT_NAME_COL     = 8;  // H
const CERT_ID_COL       = 5;  // E
const CERT_ARABIC_COL   = 7;  // G
const CERT_INCLUDE_COL  = 26; // Z — "Y" means this student gets a certificate

const CERT_CSV_HEADERS  = ['Student ID', 'Gender', 'Student Name', 'Arabic Name'];

// ─── Entry Point ───────────────────────────────────────────────────────────────

function exportCertificateCsv() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tab = ss.getSheetByName(PROGRESS_TAB);
  if (!tab) throw new Error(`Tab "${PROGRESS_TAB}" not found in the active spreadsheet.`);

  const headerRow = findHeaderRow_(tab);
  if (!headerRow) {
    throw new Error(`Could not find header "${PC_HEADER_ANCHOR}" in column ${PC_HEADER_ANCHOR_COL} ` +
                    `within the first ${PC_HEADER_SCAN_ROWS} rows of ${PROGRESS_TAB}.`);
  }

  const { rows, warnings } = collectCertificateRows_(tab, headerRow);
  if (rows.length === 0) {
    throw new Error(`No rows have "Y" in column Z of ${PROGRESS_TAB}. Nothing exported.`);
  }

  const csv      = [CERT_CSV_HEADERS].concat(rows).map(r => r.map(csvEscape_).join(',')).join('\r\n') + '\r\n';
  const filename = 'certificate_students_' +
                   Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd_HHmm') + '.csv';

  warnings.forEach(w => Logger.log('WARNING: ' + w));
  Logger.log(`Prepared ${rows.length} student(s) for download as "${filename}".`);

  showCertificateDownloadDialog_(csv, filename, rows.length, warnings);
}

// ─── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Returns the CSV rows [id, gender, name, arabicName] for every data row with Z = "Y",
 * plus warnings for rows that look incomplete. Warnings never drop a row.
 */
function collectCertificateRows_(tab, headerRow) {
  const first = headerRow + 1;
  const last  = tab.getLastRow();
  if (last < first) return { rows: [], warnings: [] };

  const values = tab.getRange(first, 1, last - first + 1, CERT_INCLUDE_COL).getValues();
  const rows = [];
  const warnings = [];

  values.forEach((r, i) => {
    if (String(r[CERT_INCLUDE_COL - 1] || '').trim().toUpperCase() !== 'Y') return;

    const sheetRow = first + i;
    const id      = String(r[CERT_ID_COL     - 1] || '').trim();
    const gender  = String(r[CERT_GENDER_COL - 1] || '').trim();
    const name    = String(r[CERT_NAME_COL   - 1] || '').trim();
    const arabic  = String(r[CERT_ARABIC_COL - 1] || '').trim();

    if (!id)     warnings.push(`Row ${sheetRow}: Student ID is blank.`);
    if (!name)   warnings.push(`Row ${sheetRow} (${id || '?'}): Student Name is blank.`);
    if (!arabic) warnings.push(`Row ${sheetRow} (${id || '?'} ${name}): Arabic Name is blank.`);
    if (!/^(male|female)$/i.test(gender)) {
      warnings.push(`Row ${sheetRow} (${id || '?'} ${name}): Gender is "${gender}" — expected Male or Female.`);
    }

    rows.push([id, gender, name, arabic]);
  });

  return { rows, warnings };
}

/** Quotes a CSV field when it contains a comma, quote or newline. */
function csvEscape_(v) {
  const s = String(v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/**
 * Opens a modeless dialog (so the execution ends right away) with a Download button.
 * The CSV is embedded in the page; the browser builds the file locally when the button is clicked.
 */
function showCertificateDownloadDialog_(csv, filename, count, warnings) {
  const esc  = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // JSON for embedding in a <script>: keep "<" out of the text so "</script>" cannot appear.
  const json = v => JSON.stringify(v).replace(/</g, '\\u003c');

  const shown = warnings.slice(0, 15).map(w => '<li>' + esc(w) + '</li>').join('');
  const more  = warnings.length > 15 ? `<li>…and ${warnings.length - 15} more (see the execution log).</li>` : '';
  const warn  = warnings.length
    ? `<p style="color:#b00020"><b>${warnings.length} warning(s)</b> — check these rows:</p><ul>${shown}${more}</ul>`
    : '';

  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    `<p><b>${count}</b> student(s) ready.</p>` +
    `<p>${esc(filename)}</p>` + warn +
    '<button id="dl" style="padding:8px 16px;font-size:14px">Download CSV</button>' +
    '</div>' +
    '<script>' +
    `const CSV = ${json(csv)}, NAME = ${json(filename)};` +
    'document.getElementById("dl").onclick = function () {' +
    '  const blob = new Blob(["\\uFEFF" + CSV], { type: "text/csv;charset=utf-8" });' +
    '  const a = document.createElement("a");' +
    '  a.href = URL.createObjectURL(blob); a.download = NAME;' +
    '  document.body.appendChild(a); a.click(); a.remove();' +
    '};' +
    '</script>';

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(warnings.length ? 420 : 200),
    'Certificate CSV');
}
