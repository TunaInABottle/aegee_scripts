
/**
 * onEdit trigger: when a cell in column A of the main data sheet is set to
 * "EB", fills the "Recipient" column (C) of every still-pending EB row with
 * an IFERROR(VLOOKUP(...)) formula that looks up the value in column H
 * against the 'EB body Vlookup' sheet.
 *
 * NOTE: Install as an installable onEdit trigger via
 * Extensions > Apps Script > Triggers (clock icon) > Add Trigger:
 *   - Function: onEdit
 *   - Event source: From spreadsheet
 *   - Event type: On edit
 */
function onEdit(e) {
  const range = e.range;

  // Only react to single-cell edits in column A
  if (range.getColumn() !== 1) return;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  const row = range.getRow();
  if (row === 1) return; // skip header

  const cell = e.value;
  if (!cell) return;

  // Only trigger when column A cell is exactly "EB"
  if (cell !== 'EB') return;

  // Only run on the main data sheet (tabular data, not query sheets)
  const sheet = e.source.getActiveSheet();
  const cellA1 = sheet.getRange(1, 1).getValue();
  if (cellA1 === '') return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'EB' && data[i][2] === '') {
      sheet.getRange(i + 2, 3).setFormula(
        `=IFERROR(VLOOKUP(TRIM(H${i + 2}),'EB body Vlookup'!A$1:B, 2, False), "NO MATCH FOUND")`
      );
    }
  }
}