/**
 * ProcessWakefern11s_AUTOPUSH  (V14)
 * ---------------------------------------------------------------------------
 * Office Script consumido por el flow
 *   "Wakefern - 11s block - Autopush V14 WHOLESALE 2 ALERTS"
 *   accion: Run_script_ProcessWakefern11s (Excel Online Business / RunScriptProd)
 *
 * IMPORTANTE
 * Este script NO viaja dentro del .zip de exportacion de Power Automate.
 * Los Office Scripts viven en el OneDrive del usuario
 * (/Documents/Office Scripts/) y el flow solo guarda un puntero:
 *
 *   scriptId = ms-officescript://onedrive_business_itemlink/01DYYVZGXRHHDY4HR32VGID4O5BXO2MOKB
 *
 * Ese id apunta al OneDrive de marin.c@pg.com. Si el flow se importa en otro
 * usuario/entorno, o si el script se renombro/movio/borro, la accion
 * Run_script_ProcessWakefern11s falla y toda la automatizacion se detiene.
 * Por eso este archivo existe: es la fuente reproducible del script.
 *
 * ENTRADA : la hoja "Alerts" del adjunto "Wholesale 2.xlsx".
 * SALIDA  : objeto con los 7 campos exactos que exige Parse_Script_Result.
 */

// ===========================================================================
// CONFIGURACION  -- lo unico que normalmente se toca
// ===========================================================================

/** Nombre de la hoja del reporte Wholesale 2 Alerts. */
const SHEET_NAME = "Alerts";

/** Se procesan solo los clientes cuyo "Sold to Name" contenga este texto. */
const CUSTOMER_MATCH = "WAKEFERN";

/** Tipo de bloqueo a procesar (columna "DB"). "11" = bloqueo de cita. */
const BLOCK_CODE = "11";

/**
 * Ventana de seleccion: se toman las ordenes cuyo Good Issue caiga
 * entre hoy y hoy + GI_WINDOW_DAYS (inclusive).
 * 1 = hoy y manana, que es la regla "11s block" actual.
 */
const GI_WINDOW_DAYS = 1;

/**
 * REGLA DE NEGOCIO -- AJUSTAR AQUI SI SU REGLA REAL ES OTRA.
 * Nuevo Good Issue = hoy + N dias habiles (sabado y domingo se saltan).
 */
const PUSH_GI_BUSINESS_DAYS_AHEAD = 2;

/** Nuevo RDD = nuevo Good Issue + N dias habiles. */
const RDD_OFFSET_BUSINESS_DAYS = 1;

/**
 * Offset horario para calcular "hoy". El flow corre en UTC pero el equipo
 * opera en Central America Standard Time (UTC-6, sin horario de verano),
 * igual que el formatDateTime/convertTimeZone del flow.
 */
const TZ_OFFSET_HOURS = -6;

/** Feriados a saltar, en ISO yyyy-MM-dd. Ej: ["2026-09-07"] */
const HOLIDAYS: string[] = [];

// ===========================================================================

/** Contrato de salida exigido por Parse_Script_Result en el flow. */
interface ScriptResult {
  rowCount: number;
  ossHtmlRows: string;
  clientHtmlRows: string;
  clientSentenceDate: string;
  todayIso: string;
  tomorrowIso: string;
  pushGiIso: string;
}

interface AlertRow {
  soNumber: string;
  poNumber: string;
  giIso: string;
}

function main(workbook: ExcelScript.Workbook): ScriptResult {
  const today = currentDateInBusinessTz();
  const todayIso = toIso(today);
  const tomorrowIso = toIso(addDays(today, 1));

  const pushGi = addBusinessDays(today, PUSH_GI_BUSINESS_DAYS_AHEAD);
  const newRdd = addBusinessDays(pushGi, RDD_OFFSET_BUSINESS_DAYS);
  const pushGiIso = toIso(pushGi);
  const newRddIso = toIso(newRdd);

  const empty: ScriptResult = {
    rowCount: 0,
    ossHtmlRows: "",
    clientHtmlRows: "",
    clientSentenceDate: formatLongDate(newRdd),
    todayIso: todayIso,
    tomorrowIso: tomorrowIso,
    pushGiIso: pushGiIso,
  };

  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) {
    // Nunca lanzar: un throw aqui rompe el flow completo. Devolver rowCount 0
    // hace que el flow entre en la rama "No_Matching_Rows" de forma limpia.
    return empty;
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return empty;
  }

  const values = usedRange.getValues();
  if (values.length < 2) {
    return empty;
  }

  // --- Mapeo de encabezados por nombre (tolerante a reordenamiento) --------
  const header = values[0];
  const colSoldToName = findColumn(header, "Sold to Name");
  const colDb = findColumn(header, "DB");
  const colDocNumber = findColumn(header, "Doc. Number");
  const colPoNumber = findColumn(header, "PO Number");
  const colGoodIssue = findColumn(header, "Good Issue");

  if (
    colSoldToName < 0 ||
    colDb < 0 ||
    colDocNumber < 0 ||
    colPoNumber < 0 ||
    colGoodIssue < 0
  ) {
    return empty;
  }

  // --- Fechas objetivo de la ventana de seleccion --------------------------
  const targetIsoDates: string[] = [];
  for (let offset = 0; offset <= GI_WINDOW_DAYS; offset++) {
    targetIsoDates.push(toIso(addDays(today, offset)));
  }

  // --- Filtrado -----------------------------------------------------------
  const matched: AlertRow[] = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const soldToName = asText(row[colSoldToName]).toUpperCase();
    if (soldToName.indexOf(CUSTOMER_MATCH) < 0) {
      continue;
    }

    // "DB" llega como texto ("11", "09", "CA") o como numero (11) segun como
    // Excel haya tipado la celda. normalizeBlock cubre ambos casos.
    if (normalizeBlock(row[colDb]) !== BLOCK_CODE) {
      continue;
    }

    const giIso = cellToIso(row[colGoodIssue]);
    if (!giIso || targetIsoDates.indexOf(giIso) < 0) {
      continue;
    }

    matched.push({
      soNumber: asText(row[colDocNumber]).trim(),
      poNumber: asText(row[colPoNumber]).trim(),
      giIso: giIso,
    });
  }

  // Orden estable: primero por GI, luego por SO, para que el correo salga
  // siempre igual ante el mismo insumo.
  matched.sort((a, b) => {
    if (a.giIso !== b.giIso) {
      return a.giIso < b.giIso ? -1 : 1;
    }
    return a.soNumber < b.soNumber ? -1 : a.soNumber > b.soNumber ? 1 : 0;
  });

  // --- Construccion del HTML ----------------------------------------------
  // escapeHtml es obligatorio: estas cadenas se inyectan en un JSON armado
  // con concat() en el flow, y un caracter suelto rompe el json().
  let ossHtmlRows = "";
  let clientHtmlRows = "";

  for (const item of matched) {
    ossHtmlRows +=
      "<tr><td>" +
      escapeHtml(item.soNumber) +
      "</td><td>" +
      pushGiIso +
      "</td></tr>";

    clientHtmlRows +=
      "<tr><td>" +
      escapeHtml(item.poNumber) +
      "</td><td>" +
      newRddIso +
      "</td></tr>";
  }

  return {
    rowCount: matched.length,
    ossHtmlRows: ossHtmlRows,
    clientHtmlRows: clientHtmlRows,
    clientSentenceDate: formatLongDate(newRdd),
    todayIso: todayIso,
    tomorrowIso: tomorrowIso,
    pushGiIso: pushGiIso,
  };
}

// ===========================================================================
// Utilidades
// ===========================================================================

/** Busca una columna por nombre exacto y, si falla, por coincidencia laxa. */
function findColumn(header: (string | number | boolean)[], name: string): number {
  const wanted = name.toLowerCase().trim();

  for (let i = 0; i < header.length; i++) {
    if (asText(header[i]).toLowerCase().trim() === wanted) {
      return i;
    }
  }

  // Fallback: ignora puntos y espacios ("Doc. Number" vs "Doc Number").
  const loose = wanted.replace(/[.\s]/g, "");
  for (let i = 0; i < header.length; i++) {
    if (asText(header[i]).toLowerCase().replace(/[.\s]/g, "") === loose) {
      return i;
    }
  }

  return -1;
}

function asText(value: string | number | boolean): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}

/**
 * Normaliza el codigo de bloqueo. Excel entrega "11" como texto pero a veces
 * como numero 11; en ese caso hay que devolver el padding a 2 digitos para
 * que "9" no deje de coincidir con "09".
 */
function normalizeBlock(value: string | number | boolean): string {
  if (typeof value === "number") {
    const asInt = Math.round(value);
    return asInt < 10 ? "0" + asInt : String(asInt);
  }
  return asText(value).trim();
}

/**
 * Convierte una celda de fecha a ISO yyyy-MM-dd.
 * getValues() devuelve fechas como serial de Excel (numero), pero si la
 * columna viene como texto tambien se soportan los formatos habituales.
 */
function cellToIso(value: string | number | boolean): string {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (typeof value === "number") {
    return toIso(excelSerialToDate(value));
  }

  const text = asText(value).trim();
  if (!text) {
    return "";
  }

  // yyyy-MM-dd (posiblemente con hora detras)
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return iso[1] + "-" + iso[2] + "-" + iso[3];
  }

  // MM/dd/yyyy
  const us = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (us) {
    return (
      us[3] + "-" + pad2(parseInt(us[1], 10)) + "-" + pad2(parseInt(us[2], 10))
    );
  }

  return "";
}

/**
 * Serial de Excel -> Date UTC.
 * Epoch 1899-12-30 compensa el bug historico del año bisiesto 1900.
 */
function excelSerialToDate(serial: number): Date {
  const wholeDays = Math.floor(serial);
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + wholeDays * 86400000);
}

/** Fecha actual desplazada a la zona horaria de operacion. */
function currentDateInBusinessTz(): Date {
  const now = new Date();
  const shifted = new Date(now.getTime() + TZ_OFFSET_HOURS * 3600000);
  return new Date(
    Date.UTC(
      shifted.getUTCFullYear(),
      shifted.getUTCMonth(),
      shifted.getUTCDate()
    )
  );
}

function addDays(date: Date, days: number): Date {
  return new Date(date.getTime() + days * 86400000);
}

/** Suma dias habiles saltando sabado, domingo y los feriados configurados. */
function addBusinessDays(date: Date, days: number): Date {
  let result = date;
  let remaining = days;

  while (remaining > 0) {
    result = addDays(result, 1);
    if (isBusinessDay(result)) {
      remaining--;
    }
  }

  return result;
}

function isBusinessDay(date: Date): boolean {
  const day = date.getUTCDay();
  if (day === 0 || day === 6) {
    return false;
  }
  return HOLIDAYS.indexOf(toIso(date)) < 0;
}

function toIso(date: Date): string {
  return (
    date.getUTCFullYear() +
    "-" +
    pad2(date.getUTCMonth() + 1) +
    "-" +
    pad2(date.getUTCDate())
  );
}

/** "Thursday, August 6, 2026" -- el texto que va en la frase al cliente. */
function formatLongDate(date: Date): string {
  const weekdays = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];
  const months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  return (
    weekdays[date.getUTCDay()] +
    ", " +
    months[date.getUTCMonth()] +
    " " +
    date.getUTCDate() +
    ", " +
    date.getUTCFullYear()
  );
}

function pad2(n: number): string {
  return n < 10 ? "0" + n : String(n);
}

/** Escapa el texto que se inyecta en HTML y, en cadena, en el JSON del flow. */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\\/g, "&#92;");
}
