// ---- test harness (appended after the script under test) ----
function wb(values: (string | number | boolean)[][]): ExcelScript.Workbook {
    const sheet: ExcelScript.Worksheet = { getUsedRange: () => ({ getValues: () => values }) };
    return { getWorksheet: (n: string) => (n === "Consolidations" ? sheet : undefined), getActiveWorksheet: () => sheet };
}
const HDR = ["DB", "Created_On", "Goods Issue Date", "RDD", "Plant", "Material Mix",
             "Sold-to Name1", "Ship-to Name1", "Sales Document", "PO",
             "Shipping Conditions", "Gross Weight", "Floor Position", "Status"];
function row(soldTo: string, shipTo: string, po: string): (string | number | boolean)[] {
    return ["X", 45000, 45001, 45002, "PLT", "MIX", soldTo, shipTo, "DOC" + po, po, "01", 1000, 10, "Open"];
}
const BLANK: (string | number | boolean)[] = ["", "", "", "", "", "", "", "", "", "", "", "", "", ""];

let failures = 0;
function check(label: string, actual: unknown, expected: unknown): void {
    const a = JSON.stringify(actual), e = JSON.stringify(expected);
    const ok = a === e;
    if (!ok) failures++;
    console.log((ok ? "  PASS " : "  FAIL ") + label + (ok ? "  = " + a : "\n        expected " + e + "\n        actual   " + a));
}
function pend(ds: DraftEmail[]): string[] {
    return ds.filter((d) => !d.matched).map((d) => (d as unknown as { shipTo?: string }).shipTo
        ?? d.subject.replace("SHIP WITH NEEDED - ", ""));
}

console.log("\nT1  duplicate contact row with a BLANK email must not erase a good contact");
{
    const v = [HDR, row("AWG - SPRINGFIELD", "AWG - GREAT LAKES DIV", "111"), BLANK];
    const contacts = JSON.stringify([
        { name: "AWG - GREAT LAKES DIV", email: "dave.scanlan@awginc.com" },
        { name: "AWG - GREAT LAKES DIV", email: "" },
    ]);
    const d = main(wb(v), contacts, "");
    check("drafts", d.length, 1);
    check("matched", d[0].matched, true);
    check("to", d[0].to, "dave.scanlan@awginc.com");
    check("pending rows created", pend(d), []);
}

console.log("\nT2  two spellings of ONE unmatched Ship-to must yield ONE clean Pending row");
{
    const v = [HDR,
        row("AWG - OKLAHOMA CITY", "AWG - Great Lakes Div", "111"), BLANK,
        row("AWG - SPRINGFIELD", "AWG - GREAT LAKES DIV", "222"), BLANK];
    const d = main(wb(v), JSON.stringify([]), "");
    check("drafts", d.length, 1);
    check("pending rows created", pend(d), ["AWG - Great Lakes Div"]);
    check("no combined name", pend(d)[0].indexOf("/") < 0, true);
}

console.log("\nT3  the Sold-to column must never be read as the Ship-to");
{
    const v = [HDR, row("AWG - OKLAHOMA CITY", "CREST FOODS", "111"), BLANK];
    const contacts = JSON.stringify([{ name: "AWG - OKLAHOMA CITY", email: "mike.bourdelais@awginc.com" }]);
    const d = main(wb(v), contacts, "");
    check("matched on Sold-to contact", d[0].matched, false);
    check("pending rows created", pend(d), ["CREST FOODS"]);
}

console.log("\nT4  matched Ship-tos sharing a recipient still fold into ONE email");
{
    const v = [HDR,
        row("AWG - OKLAHOMA CITY", "AWG - OKLAHOMA CITY", "111"), BLANK,
        row("AWG - OKLAHOMA CITY", "AWG - SPRINGFIELD", "222"), BLANK];
    const contacts = JSON.stringify([
        { name: "AWG - OKLAHOMA CITY", email: "mike.bourdelais@awginc.com" },
        { name: "AWG - SPRINGFIELD", email: "mike.bourdelais@awginc.com" },
    ]);
    const d = main(wb(v), contacts, "");
    check("drafts", d.length, 1);
    check("to", d[0].to, "mike.bourdelais@awginc.com");
    check("subject", d[0].subject, "SHIP WITH NEEDED - AWG - OKLAHOMA CITY/AWG - SPRINGFIELD");
    check("pending rows created", pend(d), []);
}

console.log("\nT5  a new unmatched Ship-to alongside matched ones gets its own Pending row");
{
    const v = [HDR,
        row("AWG - OKLAHOMA CITY", "AWG - OKLAHOMA CITY", "111"), BLANK,
        row("AWG - OKLAHOMA CITY", "AWG NEBRASKA", "222"), BLANK,
        row("AWG - SPRINGFIELD", "CREST FOODS", "333"), BLANK];
    const contacts = JSON.stringify([{ name: "awg - oklahoma  city ", email: "mike.bourdelais@awginc.com" }]);
    const d = main(wb(v), contacts, "");
    check("drafts", d.length, 3);
    check("pending rows created", pend(d).sort(), ["AWG NEBRASKA", "CREST FOODS"]);
}

console.log(failures === 0 ? "\nALL PASS\n" : "\n" + failures + " FAILURE(S)\n");
