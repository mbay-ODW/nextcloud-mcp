import JSZip from "jszip";
import { decodeXmlEntities, isPdfFile, isWordFile, pdfPages, wordToText } from "./documents";
import { isSpreadsheetFile, readSpreadsheet } from "./spreadsheet";

// ---------------------------------------------------------------------------
// Unified text extraction with page offsets, scan detection and a small cache.
//
// The whole document is read once; the tools then return only a window, a
// page range or the passages matching a query. That keeps large PDFs out of
// the agent's context.
// ---------------------------------------------------------------------------

export interface PageMark {
  n: number;
  start: number;
}

export interface Extracted {
  kind: string;
  text: string;
  pages: PageMark[];
  /** "Page", "Slide" or "Sheet" */
  label: string;
  needsOcr: boolean;
  warnings: string[];
}

const EXT_KIND: Record<string, string> = {
  ".pdf": "pdf",
  ".docx": "docx",
  ".docm": "docx",
  ".dotx": "docx",
  ".dotm": "docx",
  ".xlsx": "xlsx",
  ".xlsm": "xlsx",
  ".xltx": "xlsx",
  ".xltm": "xlsx",
  ".pptx": "pptx",
  ".odt": "odt",
  ".ods": "ods",
  ".odp": "odp",
  ".txt": "text",
  ".md": "text",
  ".csv": "text",
  ".tsv": "text",
  ".json": "text",
  ".xml": "text",
  ".log": "text",
  ".html": "text",
  ".htm": "text",
};

export const MAX_PDF_PAGES = 1000;
const MAX_SHEET_ROWS = 20000;
const MAX_SHEET_COLS = 200;
const MAX_ZIP_ENTRY_BYTES = 64 * 1024 * 1024;
/** A page with fewer visible characters counts as "no text" (scan / image only). */
const EMPTY_PAGE_CHARS = 10;

export const OCR_HINT =
  "This PDF has (almost) no text layer – it is probably a scan. Text cannot be extracted here; " +
  "put it into Paperless (OCR) and read it there.";

function extOf(filename: string): string {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i).toLowerCase() : "";
}

export function documentKind(filename: string, contentType?: string): string | null {
  if (isPdfFile(filename, contentType)) return "pdf";
  if (isWordFile(filename)) return "docx";
  if (isSpreadsheetFile(filename)) return "xlsx";
  return EXT_KIND[extOf(filename)] ?? null;
}

export function isExtractable(filename: string, contentType?: string): boolean {
  return documentKind(filename, contentType) !== null;
}

// ---------------------------------------------------------------------------
// Builders
// ---------------------------------------------------------------------------

class Paged {
  private parts: string[] = [];
  pages: PageMark[] = [];
  private length = 0;
  constructor(private header: (n: number, title?: string) => string) {}

  add(n: number, body: string, title?: string): void {
    const sep = this.parts.length ? "\n\n" : "";
    this.pages.push({ n, start: this.length + sep.length });
    const chunk = sep + this.header(n, title) + "\n" + body.replace(/^\n+|\n+$/g, "");
    this.parts.push(chunk);
    this.length += chunk.length;
  }

  get text(): string {
    return this.parts.join("");
  }
}

function tidy(text: string): string {
  return text
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function zipText(zip: JSZip, name: string): Promise<string> {
  const f = zip.file(name);
  if (!f) throw new Error(`${name} missing – not a valid document`);
  const size = (f as any)._data?.uncompressedSize ?? 0;
  if (size > MAX_ZIP_ENTRY_BYTES) throw new Error(`${name} is too large to extract (${size} bytes)`);
  return f.async("string");
}

// ---------------------------------------------------------------------------
// OpenDocument (odt, ods, odp) and PowerPoint (pptx)
// ---------------------------------------------------------------------------

/**
 * Text of an ODF XML fragment: text:p / text:h become lines, text:s spaces, text:tab tabs,
 * text:line-break line breaks. Notes and annotations are skipped.
 */
export function odfXmlToText(xml: string): string {
  const token =
    /<(\/?)(text:p|text:h|text:s|text:tab|text:line-break|office:annotation|text:note)\b([^>]*?)(\/?)>|<[^>]+>|([^<]+)/g;
  const out: string[] = [];
  let line = "";
  let skip = 0;
  let m: RegExpExecArray | null;
  while ((m = token.exec(xml)) !== null) {
    const [, closing, tag, attrs, selfClosing, text] = m;
    if (text !== undefined) {
      if (!skip) line += decodeXmlEntities(text);
      continue;
    }
    if (!tag) continue;
    const isClose = closing === "/";
    const isSelf = selfClosing === "/";
    switch (tag) {
      case "office:annotation":
      case "text:note":
        if (!isSelf) skip += isClose ? -1 : 1;
        skip = Math.max(0, skip);
        break;
      case "text:s": {
        if (!isClose && !skip) {
          const c = /text:c="(\d+)"/.exec(attrs);
          line += " ".repeat(c ? Math.min(parseInt(c[1], 10), 200) : 1);
        }
        break;
      }
      case "text:tab":
        if (!isClose && !skip) line += "\t";
        break;
      case "text:line-break":
        if (!isClose && !skip) line += "\n";
        break;
      case "text:p":
      case "text:h":
        if ((isClose || isSelf) && !skip) {
          out.push(line.replace(/\s+$/, ""));
          line = "";
        }
        break;
    }
  }
  if (line.trim()) out.push(line);
  return tidy(out.join("\n"));
}

/** One ODS sheet as tab-separated rows; LibreOffice's repeated empty padding is not expanded. */
export function odsTableToRows(tableXml: string, warnings: string[], name: string): string {
  const rows: string[] = [];
  const rowRe = /<table:table-row\b([^>]*?)(\/>|>([\s\S]*?)<\/table:table-row>)/g;
  let rm: RegExpExecArray | null;
  while ((rm = rowRe.exec(tableXml)) !== null) {
    const rowAttrs = rm[1];
    const inner = rm[3] ?? "";
    const cells: string[] = [];
    const cellRe =
      /<table:(?:covered-)?table-cell\b([^>]*?)(\/>|>([\s\S]*?)<\/table:(?:covered-)?table-cell>)/g;
    let cm: RegExpExecArray | null;
    while ((cm = cellRe.exec(inner)) !== null) {
      const rep = /table:number-columns-repeated="(\d+)"/.exec(cm[1]);
      const value = odfXmlToText(cm[3] ?? "").replace(/\n+/g, " ");
      const n = Math.min(rep ? parseInt(rep[1], 10) : 1, MAX_SHEET_COLS);
      for (let i = 0; i < n && cells.length < MAX_SHEET_COLS; i++) cells.push(value);
      if (cells.length >= MAX_SHEET_COLS) break;
    }
    const line = cells.join("\t").replace(/\t+$/, "");
    const rrep = /table:number-rows-repeated="(\d+)"/.exec(rowAttrs);
    const times = line && rrep ? parseInt(rrep[1], 10) : 1;
    for (let i = 0; i < times && rows.length < MAX_SHEET_ROWS; i++) rows.push(line);
    if (rows.length >= MAX_SHEET_ROWS) {
      warnings.push(`Sheet ${name}: only the first ${MAX_SHEET_ROWS} rows were read.`);
      break;
    }
  }
  while (rows.length && !rows[rows.length - 1]) rows.pop();
  return rows.join("\n");
}

async function odfContent(buffer: Buffer): Promise<string> {
  return zipText(await JSZip.loadAsync(buffer), "content.xml");
}

async function odtToText(buffer: Buffer): Promise<Extracted> {
  const xml = await odfContent(buffer);
  const body = /<office:text\b[^>]*>([\s\S]*)<\/office:text>/.exec(xml);
  return { kind: "odt", text: odfXmlToText(body ? body[1] : xml), pages: [], label: "Page", needsOcr: false, warnings: [] };
}

async function odsToText(buffer: Buffer): Promise<Extracted> {
  const xml = await odfContent(buffer);
  const out = new Paged((n, t) => `--- Sheet ${n}${t ? `: ${t}` : ""} ---`);
  const warnings: string[] = [];
  const tableRe = /<table:table\b([^>]*)>([\s\S]*?)<\/table:table>/g;
  let m: RegExpExecArray | null;
  let n = 0;
  while ((m = tableRe.exec(xml)) !== null) {
    n += 1;
    const name = /table:name="([^"]*)"/.exec(m[1]);
    const title = name ? decodeXmlEntities(name[1]) : undefined;
    out.add(n, odsTableToRows(m[2], warnings, title ?? String(n)), title);
  }
  return { kind: "ods", text: out.text, pages: out.pages, label: "Sheet", needsOcr: false, warnings };
}

async function odpToText(buffer: Buffer): Promise<Extracted> {
  const xml = await odfContent(buffer);
  const out = new Paged((n, t) => `--- Slide ${n}${t ? `: ${t}` : ""} ---`);
  const pageRe = /<draw:page\b([^>]*)>([\s\S]*?)<\/draw:page>/g;
  let m: RegExpExecArray | null;
  let n = 0;
  while ((m = pageRe.exec(xml)) !== null) {
    n += 1;
    const name = /draw:name="([^"]*)"/.exec(m[1]);
    out.add(n, odfXmlToText(m[2]), name ? decodeXmlEntities(name[1]) : undefined);
  }
  return { kind: "odp", text: out.text, pages: out.pages, label: "Slide", needsOcr: false, warnings: [] };
}

export function pptxSlideXmlToText(xml: string): string {
  const paras: string[] = [];
  const pRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
  let m: RegExpExecArray | null;
  while ((m = pRe.exec(xml)) !== null) {
    let t = "";
    const tRe = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>|<a:br\b[^>]*\/>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRe.exec(m[1])) !== null) t += tm[1] !== undefined ? decodeXmlEntities(tm[1]) : "\n";
    if (t.trim()) paras.push(t);
  }
  return paras.join("\n");
}

async function pptxToText(buffer: Buffer): Promise<Extracted> {
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => parseInt(/(\d+)\.xml$/.exec(a)![1], 10) - parseInt(/(\d+)\.xml$/.exec(b)![1], 10));
  const out = new Paged((n) => `--- Slide ${n} ---`);
  for (let i = 0; i < slides.length; i++) out.add(i + 1, pptxSlideXmlToText(await zipText(zip, slides[i])));
  return { kind: "pptx", text: out.text, pages: out.pages, label: "Slide", needsOcr: false, warnings: [] };
}

// ---------------------------------------------------------------------------
// PDF
// ---------------------------------------------------------------------------

async function pdfToExtracted(buffer: Buffer): Promise<Extracted> {
  const warnings: string[] = [];
  const { total, pages } = await pdfPages(buffer, MAX_PDF_PAGES);
  const out = new Paged((n) => `--- Page ${n} of ${total} ---`);
  let empty = 0;
  for (const p of pages) {
    if (p.text.replace(/\s/g, "").length < EMPTY_PAGE_CHARS) empty += 1;
    out.add(p.n, p.text);
  }
  if (total > pages.length) warnings.push(`Only the first ${pages.length} of ${total} pages were read.`);
  const needsOcr = pages.length > 0 && empty * 2 > pages.length;
  if (needsOcr) warnings.push(OCR_HINT);
  else if (empty) warnings.push(`${empty} page(s) without text (images/scans?).`);
  return { kind: "pdf", text: out.text, pages: out.pages, label: "Page", needsOcr, warnings };
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export async function extractDocument(
  buffer: Buffer,
  filename: string,
  contentType?: string
): Promise<Extracted> {
  const kind = documentKind(filename, contentType);
  switch (kind) {
    case "pdf":
      return pdfToExtracted(buffer);
    case "docx":
      return { kind, text: await wordToText(buffer), pages: [], label: "Page", needsOcr: false, warnings: [] };
    case "xlsx":
      return {
        kind,
        text: await readSpreadsheet(buffer, filename, { maxRows: MAX_SHEET_ROWS }),
        pages: [],
        label: "Sheet",
        needsOcr: false,
        warnings: [],
      };
    case "pptx":
      return pptxToText(buffer);
    case "odt":
      return odtToText(buffer);
    case "ods":
      return odsToText(buffer);
    case "odp":
      return odpToText(buffer);
    case "text":
      return { kind, text: buffer.toString("utf8"), pages: [], label: "Page", needsOcr: false, warnings: [] };
    default:
      throw new Error(`No text extraction for '${filename}'.`);
  }
}

// ---------------------------------------------------------------------------
// Cache: path + etag → extracted text (so paging does not download/parse again)
// ---------------------------------------------------------------------------

const CACHE_MAX_CHARS = 60_000_000;
const cache = new Map<string, Extracted>();
let cacheChars = 0;

export async function cachedExtract(
  key: string,
  load: () => Promise<{ buffer: Buffer; filename: string; contentType?: string }>
): Promise<{ doc: Extracted; cached: boolean }> {
  const hit = cache.get(key);
  if (hit) {
    cache.delete(key);
    cache.set(key, hit); // most recently used last
    return { doc: hit, cached: true };
  }
  const { buffer, filename, contentType } = await load();
  const doc = await extractDocument(buffer, filename, contentType);
  if (doc.text.length < CACHE_MAX_CHARS) {
    cache.set(key, doc);
    cacheChars += doc.text.length;
    for (const [k, v] of cache) {
      if (cacheChars <= CACHE_MAX_CHARS) break;
      cache.delete(k);
      cacheChars -= v.text.length;
    }
  }
  return { doc, cached: false };
}

export function clearExtractCache(): void {
  cache.clear();
  cacheChars = 0;
}

// ---------------------------------------------------------------------------
// Slicing: pages, windows, search
// ---------------------------------------------------------------------------

/** Parse "1-3,5,8-" into sorted page numbers within 1..count. */
export function parsePageSpec(spec: string, count: number): number[] {
  const pages = new Set<number>();
  for (const raw of spec.split(",")) {
    const part = raw.trim();
    if (!part) continue;
    const m = /^(\d*)\s*(-)?\s*(\d*)$/.exec(part);
    if (!m || (!m[1] && !m[3])) throw new Error(`Invalid page range '${part}' (use e.g. "1-5,8")`);
    const a = m[1] ? parseInt(m[1], 10) : 1;
    const b = m[2] ? (m[3] ? parseInt(m[3], 10) : count) : a;
    for (let p = Math.min(a, b); p <= Math.max(a, b); p++) if (p >= 1 && p <= count) pages.add(p);
  }
  return Array.from(pages).sort((x, y) => x - y);
}

/** Text of the selected pages, in order, joined by a blank line. */
export function selectPages(doc: Extracted, numbers: number[]): string {
  const ends = new Map<number, number>();
  doc.pages.forEach((p, i) => ends.set(p.n, i + 1 < doc.pages.length ? doc.pages[i + 1].start : doc.text.length));
  return numbers
    .map((n) => {
      const p = doc.pages.find((x) => x.n === n);
      return p ? doc.text.slice(p.start, ends.get(n)).replace(/\n+$/, "") : "";
    })
    .filter(Boolean)
    .join("\n\n");
}

export function pageAt(pages: PageMark[], offset: number): number | null {
  let current: number | null = null;
  for (const p of pages) {
    if (p.start > offset) break;
    current = p.n;
  }
  return current;
}

export interface Hit {
  page: number | null;
  offset: number;
  snippet: string;
}

/** Case-insensitive search for any of the words; overlapping windows are merged. */
export function findSnippets(
  doc: Extracted,
  query: string,
  context = 300,
  maxHits = 20
): { total: number; hits: Hit[] } {
  const terms = query.split(/\s+/).filter(Boolean);
  if (!terms.length) return { total: 0, hits: [] };
  const re = new RegExp(terms.map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"), "gi");
  const text = doc.text;
  const hits: Hit[] = [];
  let total = 0;
  let lastEnd = -1;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    total += 1;
    if (m[0].length === 0) re.lastIndex += 1;
    if (hits.length >= maxHits) continue;
    const start = Math.max(0, m.index - context);
    const end = Math.min(text.length, m.index + m[0].length + context);
    if (start <= lastEnd && hits.length) {
      const h = hits[hits.length - 1];
      h.snippet = text.slice(h.offset, end);
    } else {
      hits.push({ page: pageAt(doc.pages, m.index), offset: start, snippet: text.slice(start, end) });
    }
    lastEnd = end;
  }
  return { total, hits };
}
