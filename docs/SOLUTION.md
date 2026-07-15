# Moving the AWG flows (and the portfolio) into a Power Platform Solution

## Recommendation: yes

For a one-off flow, Solutions are overkill. For a **portfolio of 11 automations**
with shared connections and FTE tracking, Solutions are the right call. Everything
painful in this build — hardcoded folder/file/list/script IDs, re-mapping the 3
connections on every import, "re-pick this in the designer" — is exactly what
Solutions remove.

### What you gain

- **Connection references** — the 3 connections (Outlook `pgcustservw2`, SharePoint
  `marin.c`, Excel `marin.c`) are defined once and reused by every flow. No more
  per-import mapping.
- **Environment variables** — the site URL, folder paths, list GUIDs, Excel
  file/table IDs, Office Script ID, the AWG Complete folder ID, the CC list, the
  signature — all become variables you set **once per environment**. No more
  placeholder-bypass + re-pick.
- **ALM / portability** — export **managed** to move Dev → Test → Prod cleanly;
  versioning; a single package for the whole portfolio.
- **Governance** — one container to audit, back up, and hand off. Good for an IWS
  program that's being measured.

### Cost / caveats

- It's a **migration effort**, not a switch. Best done deliberately, a few flows at
  a time (start with AWG).
- Flows must live **inside** the solution (add existing, or rebuild).
- Some connectors don't support connection references everywhere — verify each.
- Managed solutions are read-only in the target env (edit in Dev, ship to Prod).

---

## How to convert (AWG first)

1. **Create a solution** in your Dev environment (Solutions → New solution), with a
   publisher + prefix (e.g. `awg`).
2. **Add connection references** for the 3 connectors, pointing at the existing
   connections.
3. **Add the flows** to the solution (Add existing → Cloud flow → the AWG flows),
   or rebuild them inside it.
4. **Replace hardcoded values with environment variables** (table below): create
   each env var, then swap the literal in the action for
   `@parameters('<schema_name>')` / the "Environment variable" dynamic value.
5. **Test in Dev**, then **Export** (managed) and **Import** into Prod. On import
   you set the env-var values and pick the connections **once**.

### Environment variables to extract (AWG flows)

| Env variable | Type | Value today |
|---|---|---|
| `SharePointSiteUrl` | Text | `https://pgone.sharepoint.com/sites/NACSO-RegionalVMI` |
| `TempReportFolderPath` | Text | `/Shared Documents/General/Customer Documents/AWG-VMC` |
| `TempReportFileName` | Text | `ShipWith_Temp.xlsx` |
| `ContactsFileId` | Text | (Contacts Data Base file id) |
| `ContactsTableId` | Text | (contacts table id) |
| `OfficeScriptId` | Text | (shipWithReport script id) |
| `AWGCompleteFolderId` | Text | `AAMk...AAAiDm4CAAA=` |
| `PendingListGuid` | Text | (AWG ShipWith Pending list id) |
| `SharedMailbox` | Text | `pgcustservw2.im@pg.com` |
| `CcList` | Text | `marquardt.jw@pg.com; jackson.vs@pg.com` |
| `SignatureHtml` | Text | `Regards, / pgcustservw2.im@pg.com / NA Order Management \| Regional` |

Connection references: `cr_office365_pgcustservw2`, `cr_sharepoint_marinc`,
`cr_excel_marinc`.

### Suggested order for the portfolio

1. AWG Ship-with (this one) + AWG Add-Contact — the pair we just built.
2. The AWG CON-00..04 family (they already share these connections).
3. The rest, grouped by owner/connector.

Do it once, and every future flow starts inside the solution — no more import
whack-a-mole.
