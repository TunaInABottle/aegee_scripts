
/**
 * onEdit trigger: when a value in the "Email subject" column (B) matches
 * one of the RESULTS patterns, auto-fills the "Recipient" column (C) with
 * a VLOOKUP formula against the 'EB body Vlookup' sheet.
 *
 * NOTE: Install as an installable onEdit trigger via
 * Extensions > Apps Script > Triggers (clock icon) > Add Trigger:
 *   - Function: onEdit
 *   - Event source: From spreadsheet
 *   - Event type: On edit
 */
function onEdit(e) {
  const range = e.range;

  // Only react to single-cell edits in column B (2)
  if (range.getColumn() !== 2) return;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  const row = range.getRow();
  if (row === 1) return; // skip header

  const value = e.value;
  if (!value) return;

  // Match: [Other EBs][RESULTS] Open Call ..., [NO CANDIDATES][RESULTS] Open Call ..., [CD Approval][RESULTS] Open Call ...
  if (!/^\[(Other EBs|NO CANDIDATES|CD Approval)\]\[RESULTS\] Open Call /.test(value)) return;

  const sheet = e.source.getActiveSheet();
  sheet.getRange(row, 3).setFormula(
    `=IFERROR(VLOOKUP(H${row},'EB body Vlookup'!A$1:B, 2, False), "NO MATCH FOUND")`
  );
}