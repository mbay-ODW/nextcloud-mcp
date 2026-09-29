import ExcelJS from "exceljs";

export type SpreadsheetFormat = "markdown" | "csv" | "json";

export interface SpreadsheetOptions {
  /** Sheet name or 1-based index. Omit to read all sheets. */
  sheet?: string | number;
  /** A1-style range, e.g. "A1:F50". Applied to each selected sheet. */
  range?: string;
  format?: SpreadsheetFormat;
  /** Max data rows returned per sheet (after range is applied). */
  maxRows?: number;
  /** Also output formulas next to their calculated values. */
  includeFormulas?: boolean;
}

export interface SheetSummary {
  index: number;
  name: string;
  rows: number;
  columns: number;
  state: string;
}

const XLSX_EXTENSIONS = [".xlsx", ".xlsm", ".xltx", ".xltm"];
const UNSUPPORTED_SPREADSHEET_EXTENSIONS = [".xls", ".ods", ".xlsb", ".numbers"];

function extOf(filename: string): string {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i).toLowerCase() : "";
}

export function isSpreadsheetFile(filename: string): boolean {
  return XLSX_EXTENSIONS.includes(extOf(filename));
}

export function isUnsupportedSpreadsheetFile(filename: string): boolean {
  return UNSUPPORTED_SPREADSHEET_EXTENSIONS.includes(extOf(filename));
}

// ---------------------------------------------------------------------------
// Cell value conversion
// ---------------------------------------------------------------------------

function formatDate(d: Date): string {
  if (isNaN(d.getTime())) return "";
  const iso = d.toISOString();
  // Excel dates without a time part come back as midnight UTC
  return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso.replace(".000Z", "Z");
}

type Scalar = string | number | boolean | null;

function cellValue(value: ExcelJS.CellValue, includeFormulas: boolean): Scalar {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return formatDate(value);
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  const v = value as any;
  if (Array.isArray(v.richText)) {
    return v.richText.map((r: { text: string }) => r.text).join("");
  }
  if ("formula" in v || "sharedFormula" in v) {
    const result = v.result;
    let resolved: Scalar;
    if (result === undefined || result === null) resolved = null;
    else if (result instanceof Date) resolved = formatDate(result);
    else if (typeof result === "object" && "error" in result) resolved = String(result.error);
    else resolved = result as Scalar;
    if (includeFormulas) {
      const formula = v.formula ?? v.sharedFormula;
      return `${resolved ?? ""} [=${formula}]`;
    }
    return resolved;
  }
  if ("hyperlink" in v) {
    const text = v.text;
    if (text && typeof text === "object" && Array.isArray(text.richText)) {
      return text.richText.map((r: { text: string }) => r.text).join("");
    }
    return text ?? v.hyperlink;
  }
  if ("error" in v) return String(v.error);
  return JSON.stringify(v);
}

// ---------------------------------------------------------------------------
// Range handling
// ---------------------------------------------------------------------------

function columnToNumber(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function numberToColumn(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

interface Bounds {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

function parseRange(range: string): Bounds {
  const m = range
    .trim()
    .toUpperCase()
    .match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!m) throw new Error(`Invalid range '${range}'. Expected A1 notation like "A1:F50".`);
  const c1 = columnToNumber(m[1]);
  const r1 = parseInt(m[2], 10);
  const c2 = m[3] ? columnToNumber(m[3]) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) : r1;
  return {
    top: Math.min(r1, r2),
    left: Math.min(c1, c2),
    bottom: Math.max(r1, r2),
    right: Math.max(c1, c2),
  };
}

// ---------------------------------------------------------------------------
// Sheet extraction
// ---------------------------------------------------------------------------

interface ExtractedSheet {
  name: string;
  firstRow: number;
  firstColumn: string;
  totalRows: number;
  rows: Scalar[][];
  truncated: boolean;
}

function extractSheet(
  ws: ExcelJS.Worksheet,
  opts: Required<Pick<SpreadsheetOptions, "maxRows" | "includeFormulas">> & { range?: string }
): ExtractedSheet {
  const bounds: Bounds = opts.range
    ? parseRange(opts.range)
    : { top: 1, left: 1, bottom: ws.rowCount, right: ws.columnCount };

  // Clip to the used area of the sheet
  const bottom = Math.min(bounds.bottom, ws.rowCount);
  const right = Math.min(bounds.right, ws.columnCount);

  const all: Scalar[][] = [];
  for (let r = bounds.top; r <= bottom; r++) {
    const row = ws.getRow(r);
    const values: Scalar[] = [];
    for (let c = bounds.left; c <= right; c++) {
      const cell = row.getCell(c);
      // Only the master cell of a merged range carries the value
      const value = cell.isMerged && cell.master !== cell ? null : cell.value;
      values.push(cellValue(value, opts.includeFormulas));
    }
    all.push(values);
  }

  // Drop trailing empty rows and columns
  const isEmpty = (v: Scalar) => v === null || v === "";
  while (all.length > 0 && all[all.length - 1].every(isEmpty)) all.pop();
  let width = 0;
  for (const row of all) {
    for (let i = row.length - 1; i >= 0; i--) {
      if (!isEmpty(row[i])) {
        width = Math.max(width, i + 1);
        break;
      }
    }
  }
  const trimmed = all.map((row) => row.slice(0, width));

  // maxRows counts data rows; the first row is treated as header and always kept
  const limit = opts.maxRows + 1;
  return {
    name: ws.name,
    firstRow: bounds.top,
    firstColumn: numberToColumn(bounds.left),
    totalRows: trimmed.length,
    rows: trimmed.slice(0, limit),
    truncated: trimmed.length > limit,
  };
}

// ---------------------------------------------------------------------------
// Output rendering
// ---------------------------------------------------------------------------

function toText(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(+v.toPrecision(15));
  return String(v);
}

function renderMarkdown(sheet: ExtractedSheet): string {
  if (sheet.rows.length === 0) return "_(empty)_";
  const esc = (v: Scalar) => toText(v).replace(/\|/g, "\\|").replace(/\r?\n/g, "<br>");
  const width = Math.max(...sheet.rows.map((r) => r.length), 1);
  const pad = (r: Scalar[]) => [...r, ...Array(width - r.length).fill(null)];
  const [header, ...body] = sheet.rows.map(pad);
  const lines = [
    `| ${header.map(esc).join(" | ")} |`,
    `| ${header.map(() => "---").join(" | ")} |`,
    ...body.map((r) => `| ${r.map(esc).join(" | ")} |`),
  ];
  return lines.join("\n");
}

function renderCsv(sheet: ExtractedSheet): string {
  const esc = (v: Scalar) => {
    const s = toText(v);
    return /[",\r\n;]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return sheet.rows.map((r) => r.map(esc).join(",")).join("\n");
}

function sheetHeading(sheet: ExtractedSheet): string {
  const shown = sheet.rows.length;
  const note = sheet.truncated
    ? ` — showing first ${shown} of ${sheet.totalRows} rows (use range/max_rows for more)`
    : "";
  return `## Sheet: ${sheet.name} (starts at ${sheet.firstColumn}${sheet.firstRow}, ${sheet.totalRows} rows)${note}`;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

async function loadWorkbook(buffer: Buffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  return wb;
}

export function summarizeSheets(wb: ExcelJS.Workbook): SheetSummary[] {
  return wb.worksheets.map((ws, i) => ({
    index: i + 1,
    name: ws.name,
    rows: ws.actualRowCount,
    columns: ws.actualColumnCount,
    state: ws.state,
  }));
}

export async function listSheets(buffer: Buffer): Promise<SheetSummary[]> {
  return summarizeSheets(await loadWorkbook(buffer));
}

export async function readSpreadsheet(
  buffer: Buffer,
  filename: string,
  options: SpreadsheetOptions = {}
): Promise<string> {
  const format = options.format ?? "markdown";
  const maxRows = options.maxRows ?? 500;
  const includeFormulas = options.includeFormulas ?? false;

  const wb = await loadWorkbook(buffer);
  const summaries = summarizeSheets(wb);

  let worksheets = wb.worksheets;
  if (options.sheet !== undefined && options.sheet !== "") {
    const s = options.sheet;
    const asNumber = typeof s === "number" ? s : /^\d+$/.test(s) ? parseInt(s, 10) : NaN;
    const found =
      wb.worksheets.find((ws) => ws.name === s) ??
      wb.worksheets.find((ws) => ws.name.toLowerCase() === String(s).toLowerCase()) ??
      (!isNaN(asNumber) ? wb.worksheets[asNumber - 1] : undefined);
    if (!found) {
      throw new Error(
        `Sheet '${s}' not found. Available sheets: ${summaries.map((x) => `${x.index}: ${x.name}`).join(", ")}`
      );
    }
    worksheets = [found];
  }

  const extracted = worksheets.map((ws) =>
    extractSheet(ws, { maxRows, includeFormulas, range: options.range })
  );

  if (format === "json") {
    return JSON.stringify(
      {
        file: filename,
        sheets: summaries,
        data: extracted.map((s) => ({
          sheet: s.name,
          start_cell: `${s.firstColumn}${s.firstRow}`,
          total_rows: s.totalRows,
          truncated: s.truncated,
          rows: s.rows,
        })),
      },
      null,
      2
    );
  }

  const parts: string[] = [
    `# ${filename}`,
    `Sheets: ${summaries.map((x) => `${x.index}: ${x.name}${x.state !== "visible" ? ` (${x.state})` : ""}`).join(", ")}`,
  ];
  for (const sheet of extracted) {
    parts.push("", sheetHeading(sheet), "");
    parts.push(format === "csv" ? renderCsv(sheet) : renderMarkdown(sheet));
  }
  return parts.join("\n");
}
