/**
 * FillCancellationForm
 * -------------------------------------------------------------------------
 * Office Script used by the "AWG - CAN - 02 Fill Form & Send to RPA" flow.
 *
 * Receives the orders to cancel as a single text blob (one order per line,
 * pipe-delimited: SalesOrderNumber|ReasonCode|ReasonForCancellation) and
 * writes them into the Order Cancellation Form template starting at row 3
 * (A=Sales Order Number, B=Reason Code, C=Reason for cancelation).
 *
 * The template layout is preserved (row 1 title, row 2 headers, dropdown
 * validations on B3:B82 and C3:C82). Only cell VALUES are written, so the
 * data-validation dropdowns stay intact.
 *
 * Parameter:
 *   rowsText : string  e.g. "12345|03|No longer needed\n67890|05|Duplicate order"
 * Returns:
 *   { written:number }
 * -------------------------------------------------------------------------
 */
function main(workbook: ExcelScript.Workbook, rowsText: string): { written: number } {
  const sheet = workbook.getWorksheet("Sheet1") ?? workbook.getActiveWorksheet();

  const FIRST_DATA_ROW = 3; // rows 1-2 are title/header
  const LAST_TEMPLATE_ROW = 82; // template validation range ends at row 82

  // Clear any previous data in A:C of the data area (values only).
  const clearRange = sheet.getRange(`A${FIRST_DATA_ROW}:C${LAST_TEMPLATE_ROW}`);
  clearRange.clear(ExcelScript.ClearApplyTo.contents);

  if (!rowsText) return { written: 0 };

  // Split into lines (accept \n, \r\n or \r) and drop empties/headers/markers.
  const lines = rowsText
    .split(/\r\n|\r|\n/)
    .map((l) => l.trim())
    .filter(
      (l) =>
        l.length > 0 &&
        l.indexOf("|") >= 0 &&
        l.toLowerCase().indexOf("salesordernumber") < 0 &&
        l.indexOf("###") < 0
    );

  let written = 0;
  for (let i = 0; i < lines.length; i++) {
    const parts = lines[i].split("|");
    const so = (parts[0] ?? "").trim();
    const code = (parts[1] ?? "").trim();
    const reason = (parts[2] ?? "").trim();
    if (so === "") continue;

    const r = FIRST_DATA_ROW + written;
    if (r > LAST_TEMPLATE_ROW) break; // stay within the validated template area
    sheet.getRange(`A${r}`).setValue(so);
    sheet.getRange(`B${r}`).setValue(code);
    sheet.getRange(`C${r}`).setValue(reason);
    written++;
  }

  return { written };
}
