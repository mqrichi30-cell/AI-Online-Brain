/**
 * AWG - Ship with POs Report  ·  Office Script (Excel Online)
 * ==========================================================
 * Runs from Power Automate via the "Run script" action against the report that
 * arrives attached to the "AWG - Ship with POs Repot" email (already saved to
 * OneDrive/SharePoint by the flow).
 *
 * Responsibilities
 *   1. Parse the "Consolidations" sheet into consolidation blocks.
 *   2. Match each block's Ship-to name against the Contacts Data Base (passed in
 *      by the flow as a JSON string) to resolve the recipient email.
 *   3. Group consolidations by recipient and build the draft email HTML
 *      (faithful to the original VBA `CreateConsolidationEmails`).
 *   4. Return one payload per draft. Power Automate creates the drafts.
 *
 * It intentionally does NOT send or draft anything itself — Office Scripts
 * cannot. The flow turns each returned item into a draft in the shared mailbox.
 *
 * Parameters
 *   workbook      - the report workbook (bound by the "Run script" action).
 *   contactsJson  - JSON string: [{ "name": "...", "email": "..." }, ...]
 *                   Built by the flow from the Contacts Data Base table.
 *   signatureHtml - optional HTML signature appended to every draft.
 *
 * Returns  DraftEmail[]  ->  [{ to, cc, subject, htmlBody, matched }]
 */

interface Contact {
    name: string;
    email: string;
}

interface DraftEmail {
    to: string;
    cc: string;
    subject: string;
    htmlBody: string;
    matched: boolean;
}

const SHARED_MAILBOX = "pgcustservw2.im@pg.com";
const CC_LIST = "marquardt.jw@pg.com; jackson.vs@pg.com";

const INTRO_HTML =
    "Dear Team,<br><br>" +
    "The following PO is undersized and requires a Ship With to make a full " +
    "truckload. Do you prefer to write a Ship With PO we can consolidate OR do " +
    "you prefer we add or increase product already on the PO?<br><br>";

// Generic default signature (used when the flow passes an empty signatureHtml).
const DEFAULT_SIGNATURE =
    "<br>Regards,<br>" +
    "pgcustservw2.im@pg.com<br>" +
    "NA Order Management | Regional<br>";

function main(
    workbook: ExcelScript.Workbook,
    contactsJson: string,
    signatureHtml: string = ""
): DraftEmail[] {
    const sheet =
        workbook.getWorksheet("Consolidations") || workbook.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    if (!usedRange) {
        return [];
    }

    const values = usedRange.getValues();
    if (values.length < 2) {
        return [];
    }

    // --- Locate columns by header (tolerant to reordering) --------------------
    const header = values[0].map((h) => String(h).trim().toLowerCase());
    const col = (...aliases: string[]): number => {
        for (const a of aliases) {
            const idx = header.findIndex((h) => h === a.toLowerCase());
            if (idx >= 0) return idx;
        }
        // fallback: partial contains
        for (const a of aliases) {
            const idx = header.findIndex((h) => h.indexOf(a.toLowerCase()) >= 0);
            if (idx >= 0) return idx;
        }
        return -1;
    };

    const cDb = col("db", "delivery block");
    const cCreated = col("created_on", "created on");
    const cGi = col("goods issue date", "gi");
    const cRdd = col("rdd", "requested delivery date");
    const cPlant = col("plant");
    const cMix = col("material mix");
    const cShipTo = col("ship-to name1", "ship to name", "ship-to name");
    const cDoc = col("sales document", "document number");
    const cPo = col("po", "purchase order number");
    const cShipCond = col("shipping conditions", "shipping condition");
    const cWeight = col("gross weight");
    const cFp = col("floor position", "order header fp", "fp");
    const cStatus = col("status", "description");

    // --- Parse consolidation blocks -------------------------------------------
    const blocks: ConsolidationBlock[] = [];
    let current: (string | number | boolean)[][] = [];
    let number = 0;

    const isDataRow = (r: (string | number | boolean)[]): boolean =>
        String(r[cDb] ?? "").trim() !== "" &&
        String(r[cShipTo] ?? "").trim() !== "";

    for (let i = 1; i < values.length; i++) {
        const r = values[i];
        if (isDataRow(r)) {
            current.push(r);
        } else if (current.length > 0) {
            number++;
            blocks.push(buildBlock(number, current, {
                cCreated, cGi, cRdd, cPlant, cMix, cShipTo, cDoc, cPo, cShipCond, cWeight, cFp, cStatus,
            }));
            current = [];
        }
    }
    if (current.length > 0) {
        number++;
        blocks.push(buildBlock(number, current, {
            cCreated, cGi, cRdd, cPlant, cMix, cShipTo, cDoc, cPo, cShipCond, cWeight, cFp, cStatus,
        }));
    }

    // --- Contact lookup --------------------------------------------------------
    const contacts: Contact[] = contactsJson ? JSON.parse(contactsJson) : [];
    const contactMap: { [key: string]: string } = {};
    for (const c of contacts) {
        if (c && c.name) {
            contactMap[normName(c.name)] = String(c.email || "").trim();
        }
    }

    // --- Group blocks by resolved recipient -----------------------------------
    interface Group {
        email: string;
        matched: boolean;
        names: string[];
        blocks: ConsolidationBlock[];
    }
    const groups: { [key: string]: Group } = {};

    for (const b of blocks) {
        const resolved = contactMap[normName(b.shipTo)] || "";
        const matched = resolved !== "";
        // Unmatched Ship-tos are routed to the shared mailbox (with a note) so a
        // human can add the contact and forward — nothing is lost.
        const email = matched ? resolved : SHARED_MAILBOX;
        const key = matched ? email.toLowerCase() : "__unmatched__" + normName(b.shipTo);
        if (!groups[key]) {
            groups[key] = { email: email, matched: matched, names: [], blocks: [] };
        }
        if (groups[key].names.indexOf(b.shipTo) < 0) {
            groups[key].names.push(b.shipTo);
        }
        groups[key].blocks.push(b);
    }

    // --- Build one draft payload per group ------------------------------------
    const drafts: DraftEmail[] = [];
    for (const key of Object.keys(groups)) {
        const g = groups[key];
        // For unmatched Ship-tos, add a "please add contact" note at the very top.
        let body = "";
        if (!g.matched) {
            body +=
                "<p style='color:#b00020;font-weight:bold;'>Please add contact for " +
                esc(g.names.join(", ")) + ".</p>";
        }
        body += INTRO_HTML;
        for (const b of g.blocks) {
            body += blockTableHtml(b);
        }
        body += signatureHtml && signatureHtml.length > 0 ? signatureHtml : DEFAULT_SIGNATURE;
        drafts.push({
            to: g.email,
            cc: CC_LIST,
            subject: "SHIP WITH NEEDED - " + g.names.join("/"),
            htmlBody: "<html><body>" + body + "</body></html>",
            matched: g.matched,
        });
    }

    return drafts;
}

// ---------------------------------------------------------------------------

interface RowData {
    createdOn: string;
    gi: string;
    rdd: string;
    plant: string;
    mix: string;
    shipTo: string;
    doc: string;
    po: string;
    shipCond: string;
    weight: number;
    fp: number;
    status: string;
}

interface ConsolidationBlock {
    number: number;
    shipTo: string;
    rows: RowData[];
}

interface ColMap {
    cCreated: number; cGi: number; cRdd: number; cPlant: number; cMix: number;
    cShipTo: number; cDoc: number; cPo: number; cShipCond: number;
    cWeight: number; cFp: number; cStatus: number;
}

function buildBlock(
    num: number,
    rows: (string | number | boolean)[][],
    c: ColMap
): ConsolidationBlock {
    return {
        number: num,
        shipTo: String(rows[0][c.cShipTo] ?? ""),
        rows: rows.map((r) => ({
            createdOn: fmtDate(r[c.cCreated]),
            gi: fmtDate(r[c.cGi]),
            rdd: fmtDate(r[c.cRdd]),
            plant: String(r[c.cPlant] ?? ""),
            mix: String(r[c.cMix] ?? ""),
            shipTo: String(r[c.cShipTo] ?? ""),
            doc: String(r[c.cDoc] ?? ""),
            po: String(r[c.cPo] ?? ""),
            shipCond: String(r[c.cShipCond] ?? ""),
            weight: toNum(r[c.cWeight]),
            fp: toNum(r[c.cFp]),
            status: String(r[c.cStatus] ?? ""),
        })),
    };
}

function blockTableHtml(b: ConsolidationBlock): string {
    let totalWeight = 0;
    let totalFp = 0;
    let rowsHtml = "";
    for (const r of b.rows) {
        totalWeight += r.weight;
        totalFp += r.fp;
        rowsHtml +=
            "<tr>" +
            td(r.createdOn) + td(r.gi) + td(r.rdd) + td(r.plant) + td(r.mix) +
            td(r.shipTo) + td(r.doc) + td(r.po) + td(r.shipCond) +
            td(fmtNum(r.weight, 3)) + td(fmtNum(r.fp, 3)) + td(r.status) +
            "</tr>";
    }

    const head =
        "<tr>" +
        th("Created On", true) + th("Goods Issue Date", true) +
        th("Request Delivery Day", true) + th("Plant", true) +
        th("Material Mix") + th("Ship to Name") + th("Document Number") +
        th("Purchase Order Number") + th("Shipping Condition") +
        th("Gross Weight", true) + th("FP", true) + th("Description") +
        "</tr>";

    const totals =
        "<tr>" +
        "<td colspan='9'><strong>Totals:</strong></td>" +
        "<td><strong>" + fmtNum(totalWeight, 2) + "</strong></td>" +
        "<td><strong>" + fmtNum(totalFp, 2) + "</strong></td>" +
        "<td></td></tr>";

    return (
        "<h3>Consolidation Number: " + b.number + "</h3>" +
        "<table border='1' cellpadding='5' cellspacing='0'>" +
        head + rowsHtml + totals +
        "</table><br>"
    );
}

// --- helpers ---------------------------------------------------------------

function normName(value: string): string {
    return String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
}

function toNum(value: string | number | boolean): number {
    if (value === null || value === undefined || value === "") return 0;
    const n = Number(value);
    return isNaN(n) ? 0 : n;
}

/** Excel getValues() returns dates as serial numbers; convert to M/D/YYYY. */
function fmtDate(value: string | number | boolean): string {
    if (value === null || value === undefined || value === "") return "";
    if (typeof value === "number") {
        // Excel serial date -> JS Date (accounts for the 1900 leap-year bug via 25569).
        const ms = Math.round((value - 25569) * 86400 * 1000);
        const d = new Date(ms);
        if (!isNaN(d.getTime())) {
            return (d.getUTCMonth() + 1) + "/" + d.getUTCDate() + "/" + d.getUTCFullYear();
        }
    }
    return String(value);
}

function fmtNum(value: number, decimals: number): string {
    return value.toLocaleString("en-US", {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
    });
}

function esc(value: string): string {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

function td(value: string): string {
    return "<td>" + esc(value) + "</td>";
}

function th(label: string, yellow: boolean = false): string {
    return yellow
        ? "<th style='background-color:yellow;'>" + esc(label) + "</th>"
        : "<th>" + esc(label) + "</th>";
}
