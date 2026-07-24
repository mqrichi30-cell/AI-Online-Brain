/**
 * ProcessWakefernYV_AUTOPUSH
 * -------------------------------------------------------------------------
 * Office Script for the "Wakefern - YV block - Stage 1 Request Qty" flow.
 *
 * What it does:
 *  - Reads the Regional Report sheet.
 *  - Keeps only rows where  "Del Blk Indicator" == "YV"
 *    AND the RDD date is between today and today+10 days (inclusive).
 *    (Rows whose RDD is further out than today+10 are ignored on purpose.)
 *  - Returns rowCount + an HTML fragment (ossHtmlRows) with one
 *    <tr><td>SO#</td><td>Line#</td><td>RDD</td></tr> per matching row,
 *    which the flow drops into the OSSGENAI request email.
 *
 * Return shape (must match the Parse_YV_Script_Result schema in the flow):
 *  { rowCount:number, ossHtmlRows:string, todayIso:string, maxRddIso:string }
 * -------------------------------------------------------------------------
 * SETUP: adjust the header names in COLS below to match your report exactly,
 * then Save. Copy the resulting Office Script id into the flow action
 * "Run_script_ProcessWakefernYV" (replace PASTE_YV_OFFICE_SCRIPT_ID_HERE).
 */

function main(workbook: ExcelScript.Workbook): {
  rowCount: number;
  ossHtmlRows: string;
  todayIso: string;
  maxRddIso: string;
} {
  // ---- Configuration: exact header texts in the report (case-insensitive) ----
  const COLS = {
    block: "Del Blk Indicator", // the YV block column (given by the business)
    so: "SO #",                 // Sales Order number  -> adjust if named differently
    line: "Line #",             // Order line number   -> adjust if named differently
    rdd: "RDD",                 // Requested Delivery Date -> adjust if needed
  };
  const BLOCK_VALUE = "YV";
  const MAX_DAYS = 10; // today .. today+10 inclusive

  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return { rowCount: 0, ossHtmlRows: "", todayIso: "", maxRddIso: "" };
  }

  const values = usedRange.getValues();
  if (values.length < 2) {
    return { rowCount: 0, ossHtmlRows: "", todayIso: "", maxRddIso: "" };
  }

  // ---- Locate header columns (case-insensitive, trimmed) ----
  const header = values[0].map((h) => String(h).trim().toLowerCase());
  const findCol = (name: string): number =>
    header.indexOf(name.trim().toLowerCase());

  const cBlock = findCol(COLS.block);
  const cSo = findCol(COLS.so);
  const cLine = findCol(COLS.line);
  const cRdd = findCol(COLS.rdd);

  if (cBlock < 0 || cRdd < 0) {
    // Header not found -> nothing we can safely do.
    return { rowCount: 0, ossHtmlRows: "", todayIso: "", maxRddIso: "" };
  }

  // ---- Date window: today .. today+10 (date-only, no time component) ----
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const maxRdd = new Date(today.getFullYear(), today.getMonth(), today.getDate() + MAX_DAYS);

  const toIso = (d: Date): string => {
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${d.getFullYear()}-${m}-${day}`;
  };

  // Excel serial date (days since 1899-12-30) -> JS Date (date only)
  const excelSerialToDate = (serial: number): Date => {
    const ms = Math.round((serial - 25569) * 86400 * 1000); // 25569 = days to 1970-01-01
    const d = new Date(ms);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  };

  const parseRdd = (raw: string | number | boolean): Date | null => {
    if (raw === null || raw === undefined || raw === "") return null;
    if (typeof raw === "number") return excelSerialToDate(raw);
    const s = String(raw).trim();
    const parsed = new Date(s);
    if (isNaN(parsed.getTime())) return null;
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  };

  const esc = (v: string | number | boolean): string =>
    String(v)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");

  // ---- Filter rows ----
  const rows: string[] = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const blockVal = String(row[cBlock]).trim().toUpperCase();
    if (blockVal !== BLOCK_VALUE) continue;

    const rdd = parseRdd(row[cRdd]);
    if (!rdd) continue; // no valid date -> do not work it
    if (rdd < today || rdd > maxRdd) continue; // outside today..today+10 -> skip

    const so = cSo >= 0 ? esc(row[cSo]) : "";
    const line = cLine >= 0 ? esc(row[cLine]) : "";
    rows.push(
      `<tr><td>${so}</td><td>${line}</td><td>${toIso(rdd)}</td></tr>`
    );
  }

  return {
    rowCount: rows.length,
    ossHtmlRows: rows.join(""),
    todayIso: toIso(today),
    maxRddIso: toIso(maxRdd),
  };
}
