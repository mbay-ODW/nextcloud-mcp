import JSZip from "jszip";

// ---------------------------------------------------------------------------
// File type detection
// ---------------------------------------------------------------------------

const WORD_EXTENSIONS = [".docx", ".docm", ".dotx", ".dotm"];
const PDF_EXTENSIONS = [".pdf"];
/** ZIP containers whose entries can be listed and read (e.g. EVEBI project files .evex/.evea). */
const ARCHIVE_EXTENSIONS = [".zip", ".evex", ".evea", ".odt", ".ods", ".odp", ".epub"];
const TEXT_ENTRY_EXTENSIONS = [
  ".xml", ".txt", ".json", ".csv", ".html", ".htm", ".md", ".yaml", ".yml", ".ini", ".rels", ".svg",
];

function extOf(filename: string): string {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i).toLowerCase() : "";
}

export function isWordFile(filename: string): boolean {
  return WORD_EXTENSIONS.includes(extOf(filename));
}

export function isPdfFile(filename: string, contentType?: string): boolean {
  return PDF_EXTENSIONS.includes(extOf(filename)) || contentType === "application/pdf";
}

export function isArchiveFile(filename: string): boolean {
  return ARCHIVE_EXTENSIONS.includes(extOf(filename));
}

export function isDocumentFile(filename: string, contentType?: string): boolean {
  return isWordFile(filename) || isPdfFile(filename, contentType);
}

// ---------------------------------------------------------------------------
// Paging of long texts
// ---------------------------------------------------------------------------

export const DEFAULT_MAX_CHARS = 100_000;

/** Cut a long text into a window and say how to get the rest. */
export function page(text: string, offset = 0, maxChars = DEFAULT_MAX_CHARS): string {
  const total = text.length;
  const start = Math.max(0, Math.min(offset, total));
  const end = Math.min(total, start + maxChars);
  const chunk = text.slice(start, end);
  if (start === 0 && end === total) return chunk;
  const note =
    end < total
      ? `\n\n[Characters ${start}–${end} of ${total}. Continue with offset=${end}.]`
      : `\n\n[Characters ${start}–${end} of ${total}. End of text.]`;
  return chunk + note;
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

const ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

export function decodeXmlEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e: string) => {
    if (e[0] === "#") {
      const code = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e] ?? m;
  });
}

// ---------------------------------------------------------------------------
// Word (.docx)
// ---------------------------------------------------------------------------

/**
 * Text of a WordprocessingML part (document.xml, header, footer).
 *
 * Paragraphs become lines; tables become one line per row with the cells separated by tabs (several
 * paragraphs within one cell are joined with " / "), so tables stay readable and machine-parsable.
 */
export function wordXmlToText(xml: string): string {
  const token =
    /<(\/?)(w:p|w:tc|w:tr|w:tbl|w:t|w:tab|w:br|w:cr|w:noBreakHyphen|w:softHyphen)\b[^>]*?(\/?)>|<[^>]+>|([^<]+)/g;
  const out: string[] = [];
  let line = "";
  let cell = "";
  let cellParas: string[] = [];
  let row: string[] = [];
  let tcDepth = 0;
  let inText = false;
  let m: RegExpExecArray | null;

  const write = (s: string) => {
    if (tcDepth > 0) cell += s;
    else line += s;
  };

  while ((m = token.exec(xml)) !== null) {
    const [, closing, tag, selfClosing, text] = m;
    if (text !== undefined) {
      if (inText) write(decodeXmlEntities(text));
      continue;
    }
    if (!tag) continue;
    const isClose = closing === "/";
    const isSelf = selfClosing === "/";
    switch (tag) {
      case "w:t":
        if (!isSelf) inText = !isClose;
        break;
      case "w:tab":
        if (!isClose) write("\t");
        break;
      case "w:br":
      case "w:cr":
        if (!isClose) write(tcDepth > 0 ? " " : "\n");
        break;
      case "w:noBreakHyphen":
        if (!isClose) write("-");
        break;
      case "w:p":
        if (isClose) {
          if (tcDepth > 0) {
            if (cell.trim()) cellParas.push(cell.trim());
            cell = "";
          } else {
            out.push(line.replace(/\s+$/, ""));
            line = "";
          }
        }
        break;
      case "w:tc":
        if (isSelf) {
          row.push("");
        } else if (!isClose) {
          tcDepth += 1;
          if (tcDepth === 1) {
            cell = "";
            cellParas = [];
          }
        } else {
          if (tcDepth === 1) {
            if (cell.trim()) cellParas.push(cell.trim());
            row.push(cellParas.join(" / ").replace(/\t/g, " "));
            cell = "";
            cellParas = [];
          }
          tcDepth = Math.max(0, tcDepth - 1);
        }
        break;
      case "w:tr":
        if (isClose && tcDepth === 0) {
          out.push(row.join("\t").replace(/\s+$/, ""));
          row = [];
        } else if (!isClose && tcDepth === 0) {
          row = [];
        }
        break;
      case "w:tbl":
        if (tcDepth === 0) out.push("");
        break;
    }
  }
  if (line.trim()) out.push(line);
  return out
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export async function wordToText(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const main = zip.file("word/document.xml");
  if (!main) throw new Error("Not a Word document (word/document.xml missing)");
  const body = wordXmlToText(await main.async("string"));
  const parts: string[] = [];
  const headers = Object.keys(zip.files)
    .filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n))
    .sort();
  for (const name of headers) {
    const t = wordXmlToText(await zip.file(name)!.async("string"));
    if (t) parts.push(`[${name.slice(5, -4)}] ${t.replace(/\n+/g, " | ")}`);
  }
  return parts.length ? `${body}\n\n${parts.join("\n")}` : body;
}

// ---------------------------------------------------------------------------
// PDF
// ---------------------------------------------------------------------------

// pdfjs-dist is an ES module; load it with a real dynamic import (TypeScript would turn import() into require()).
const importEsm = new Function("specifier", "return import(specifier)") as (s: string) => Promise<any>;

/** Parse "1-3,5" into a sorted list of 1-based page numbers within 1..count. */
export function parsePages(spec: string | undefined, count: number): number[] {
  if (!spec || !spec.trim()) return Array.from({ length: count }, (_, i) => i + 1);
  const pages = new Set<number>();
  for (const part of spec.split(",")) {
    const m = part.trim().match(/^(\d+)\s*(?:-\s*(\d+))?$/);
    if (!m) throw new Error(`Invalid page range '${part.trim()}' (use e.g. "1-5,8")`);
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    for (let p = Math.min(a, b); p <= Math.max(a, b); p++) if (p >= 1 && p <= count) pages.add(p);
  }
  return [...pages].sort((x, y) => x - y);
}

/** Text of a PDF, one block per page ("--- Page n of N ---"); lines follow the PDF's end-of-line marks. */
export async function pdfToText(buffer: Buffer, pages?: string): Promise<string> {
  const pdfjs = await importEsm("pdfjs-dist/legacy/build/pdf.mjs");
  const task = pdfjs.getDocument({
    data: new Uint8Array(buffer),
    isEvalSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
    verbosity: 0,
  });
  const doc = await task.promise;
  try {
    const out: string[] = [];
    for (const n of parsePages(pages, doc.numPages)) {
      const p = await doc.getPage(n);
      const content = await p.getTextContent();
      let text = "";
      for (const item of content.items as any[]) {
        if (typeof item.str !== "string") continue;
        text += item.str;
        text += item.hasEOL ? "\n" : item.str && !item.str.endsWith(" ") ? " " : "";
      }
      out.push(`--- Page ${n} of ${doc.numPages} ---\n${text.replace(/[ \t]+\n/g, "\n").trim()}`);
      p.cleanup();
    }
    return out.join("\n\n");
  } finally {
    await doc.destroy();
  }
}

// ---------------------------------------------------------------------------
// ZIP archives (e.g. EVEBI .evex)
// ---------------------------------------------------------------------------

export interface ArchiveEntry {
  name: string;
  size: number;
  compressed_size: number;
  is_directory: boolean;
}

export async function listArchive(buffer: Buffer): Promise<ArchiveEntry[]> {
  const zip = await JSZip.loadAsync(buffer);
  return Object.values(zip.files).map((f: any) => ({
    name: f.name,
    size: f._data?.uncompressedSize ?? 0,
    compressed_size: f._data?.compressedSize ?? 0,
    is_directory: f.dir,
  }));
}

export function isTextEntry(name: string): boolean {
  return TEXT_ENTRY_EXTENSIONS.includes(extOf(name)) || !name.includes(".");
}

/** Text of one archive entry (XML, text). Word, PDF and nested archives are converted like files. */
export async function readArchiveEntry(buffer: Buffer, entry: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file(entry);
  if (!f) {
    const names = Object.keys(zip.files).slice(0, 50).join(", ");
    throw new Error(`Entry '${entry}' not found. Entries: ${names}`);
  }
  if (isWordFile(entry)) return wordToText(await f.async("nodebuffer"));
  if (isPdfFile(entry)) return pdfToText(await f.async("nodebuffer"));
  const data = await f.async("nodebuffer");
  if (!isTextEntry(entry) && data.subarray(0, 4096).includes(0)) {
    throw new Error(`Entry '${entry}' is binary (${data.length} bytes) and cannot be returned as text.`);
  }
  return data.toString("utf8");
}

/** Text of a Word or PDF file. */
export async function documentToText(buffer: Buffer, filename: string, pages?: string): Promise<string> {
  if (isWordFile(filename)) return wordToText(buffer);
  return pdfToText(buffer, pages);
}
