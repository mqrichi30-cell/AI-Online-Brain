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
 * Contact identity
 *   A contact is keyed on the NORMALISED SHIP-TO NAME and nothing else. The
 *   Sold-to never participates: it is excluded from header resolution so it can
 *   never be read as a Ship-to, and it is never part of a grouping key. Every
 *   unmatched Ship-to yields exactly one draft, so the flow creates exactly one
 *   Pending row per Ship-to and `shipTo` below is always a single clean name.
 *
 * Parameters
 *   workbook      - the report workbook (bound by the "Run script" action).
 *   contactsJson  - JSON string: [{ "name": "...", "email": "..." }, ...]
 *                   Built by the flow from the Contacts Data Base table.
 *   signatureHtml - optional HTML signature appended to every draft.
 *
 * Returns  DraftEmail[]  ->  [{ to, cc, subject, shipTo, shipTos, htmlBody, matched }]
 *   `shipTo` is the value the flow must write into the Contacts Data Base when
 *   `matched` is false. Do NOT derive it from `subject`.
 */

interface Contact {
    name: string;
    email: string;
}

interface DraftEmail {
    to: string;
    cc: string;
    subject: string;
    /** Canonical Ship-to for this draft. Exactly one name when matched === false. */
    shipTo: string;
    /** Every Ship-to folded into this draft (length 1 when matched === false). */
    shipTos: string[];
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

    /**
     * Resolves a column index by header.
     *   - exact alias match first, then partial (contains) match;
     *   - `blocked` header indexes are never returned, so a Sold-to column can
     *     never be picked up by the Ship-to lookup's partial pass.
     */
    const col = (aliases: string[], blocked: number[] = []): number => {
        const usable = (idx: number): boolean => idx >= 0 && blocked.indexOf(idx) < 0;
        for (const a of aliases) {
            const needle = a.toLowerCase();
            for (let i = 0; i < header.length; i++) {
                if (header[i] === needle && usable(i)) return i;
            }
        }
        for (const a of aliases) {
            const needle = a.toLowerCase();
            for (let i = 0; i < header.length; i++) {
                if (header[i].indexOf(needle) >= 0 && usable(i)) return i;
            }
        }
        return -1;
    };

    // Resolve the Sold-to FIRST and block it, so the Ship-to lookup can never
    // land on it. This is what used to let a Sold-to name be stored as a contact.
    const cSoldTo = col(["sold-to name1", "sold to name", "sold-to name", "sold-to", "sold to"]);
    const soldToBlocked: number[] = cSoldTo >= 0 ? [cSoldTo] : [];

    const cDb = col(["db", "delivery block"]);
    const cCreated = col(["created_on", "created on"]);
    const cGi = col(["goods issue date", "gi"]);
    const cRdd = col(["rdd", "requested delivery date"]);
    const cPlant = col(["plant"]);
    const cMix = col(["material mix"]);
    const cShipTo = col(["ship-to name1", "ship to name", "ship-to name", "ship-to"], soldToBlocked);
    const cDoc = col(["sales document", "document number"]);
    const cPo = col(["po", "purchase order number"]);
    const cShipCond = col(["shipping conditions", "shipping condition"]);
    const cWeight = col(["gross weight"]);
    const cFp = col(["floor position", "order header fp", "fp"]);
    const cStatus = col(["status", "description"]);

    // Without a Ship-to column there is no contact identity — bail out rather
    // than silently grouping everything under one bogus key.
    if (cShipTo < 0) {
        return [];
    }

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
    // Keyed on the normalised Ship-to name. Rows with a blank email are skipped
    // and the first non-blank email wins, so a duplicate or half-filled row can
    // never wipe out a Ship-to that already has a contact.
    const contacts: Contact[] = contactsJson ? JSON.parse(contactsJson) : [];
    const contactMap: { [key: string]: string } = {};
    for (const c of contacts) {
        if (!c || !c.name) continue;
        const email = String(c.email || "").trim();
        if (email === "") continue;
        const key = normName(c.name);
        if (contactMap[key] === undefined) {
            contactMap[key] = email;
        }
    }

    // --- Fold blocks into one entry per Ship-to --------------------------------
    // The Ship-to is the only identity. Two spellings of the same Ship-to
    // ("AWG - Great Lakes Div" / "AWG - GREAT LAKES DIV") collapse into a single
    // entry that carries one canonical display name.
    interface ShipToEntry {
        key: string;
        name: string;
        email: string;
        matched: boolean;
        blocks: ConsolidationBlock[];
    }
    const entries: { [key: string]: ShipToEntry } = {};
    const entryOrder: string[] = [];

    for (const b of blocks) {
        const key = normName(b.shipTo);
        if (key === "") continue;
        if (!entries[key]) {
            const resolved = contactMap[key] || "";
            entries[key] = {
                key: key,
                // Canonical display name: the first spelling seen for this Ship-to.
                name: String(b.shipTo).replace(/\s+/g, " ").trim(),
                email: resolved,
                matched: resolved !== "",
                blocks: [],
            };
            entryOrder.push(key);
        }
        entries[key].blocks.push(b);
    }

    // --- Group into drafts -----------------------------------------------------
    // Matched Ship-tos that share a recipient are folded into one email, so a
    // person still receives a single message. Unmatched Ship-tos are NEVER
    // folded: one draft per Ship-to keeps the Pending row (and therefore the
    // Contacts Data Base) keyed on exactly one clean Ship-to name.
    interface Group {
        email: string;
        matched: boolean;
        names: string[];
        blocks: ConsolidationBlock[];
    }
    const groups: { [key: string]: Group } = {};
    const groupOrder: string[] = [];

    for (const k of entryOrder) {
        const e = entries[k];
        const groupKey = e.matched ? "m:" + e.email.toLowerCase() : "u:" + e.key;
        const to = e.matched ? e.email : SHARED_MAILBOX;
        if (!groups[groupKey]) {
            groups[groupKey] = { email: to, matched: e.matched, names: [], blocks: [] };
            groupOrder.push(groupKey);
        }
        if (groups[groupKey].names.indexOf(e.name) < 0) {
            groups[groupKey].names.push(e.name);
        }
        for (const b of e.blocks) {
            groups[groupKey].blocks.push(b);
        }
    }

    // --- Build one draft payload per group ------------------------------------
    const drafts: DraftEmail[] = [];
    for (const groupKey of groupOrder) {
        const g = groups[groupKey];
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
            shipTo: g.names.join("/"),
            shipTos: g.names,
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
