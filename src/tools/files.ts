// @ts-nocheck
import { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { NextcloudAPI } from "../api/NextcloudAPI";
import {
  isSpreadsheetFile,
  isUnsupportedSpreadsheetFile,
  readSpreadsheet,
} from "../utils/spreadsheet";
import {
  DEFAULT_MAX_CHARS,
  isArchiveFile,
  isDocumentFile,
  listArchive,
  page,
  readArchiveEntry,
} from "../utils/documents";
import {
  cachedExtract,
  findSnippets,
  isExtractable,
  parsePageSpec,
  selectPages,
} from "../utils/extract";

const MAX_SPREADSHEET_BYTES = 50 * 1024 * 1024;
const MAX_DOCUMENT_BYTES = 80 * 1024 * 1024;

const pagingSchema = {
  offset: z
    .number()
    .int()
    .min(0)
    .optional()
    .default(0)
    .describe("Character offset to start from (for long texts; the reply says where to continue)"),
  max_chars: z
    .number()
    .int()
    .min(1000)
    .max(400000)
    .optional()
    .default(DEFAULT_MAX_CHARS)
    .describe(`Maximum characters returned (default ${DEFAULT_MAX_CHARS})`),
};

/**
 * Read a document via the extraction cache and return the requested part:
 * the passages for a query, selected pages, or a window of the running text.
 */
async function readExtracted(
  api: NextcloudAPI,
  path: string,
  info: any,
  opts: { pages?: string; query?: string; offset: number; max_chars: number }
) {
  const key = `${info.path ?? path}|${info.etag ?? ""}|${info.size}`;
  const { doc } = await cachedExtract(key, async () => ({
    buffer: await api.getFileBuffer(path),
    filename: info.name,
    contentType: info.content_type,
  }));
  const label = doc.label.toLowerCase();
  const pageInfo = doc.pages.length ? `, ${doc.pages.length} ${label}s` : "";
  const header = `[${info.name}: ${doc.kind.toUpperCase()}${pageInfo}, ${doc.text.length} characters]`;
  const warnings = doc.warnings.map((w) => `[Warning: ${w}]`).join("\n");
  const head = warnings ? `${header}\n${warnings}` : header;

  if (opts.query) {
    const { total, hits } = findSnippets(doc, opts.query);
    const result: any = {
      file: info.name,
      kind: doc.kind,
      page_count: doc.pages.length || null,
      text_length: doc.text.length,
      needs_ocr: doc.needsOcr,
      query: opts.query,
      total_hits: total,
      hits,
    };
    if (doc.warnings.length) result.warnings = doc.warnings;
    if (total > hits.length)
      result.note = "More hits than shown – narrow the query or read by pages.";
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }

  let text = doc.text;
  if (opts.pages) {
    if (!doc.pages.length) {
      return {
        content: [
          {
            type: "text",
            text: `${head}\n[Note: this file type has no ${label}s; 'pages' was ignored.]\n\n${page(text, opts.offset, opts.max_chars)}`,
          },
        ],
      };
    }
    const wanted = parsePageSpec(opts.pages, doc.pages.length);
    if (!wanted.length) {
      return {
        content: [
          {
            type: "text",
            text: `Error: no such ${label}. ${doc.label}s 1–${doc.pages.length} exist.`,
          },
        ],
        isError: true,
      };
    }
    text = selectPages(doc, wanted);
  }
  const body = page(text, opts.offset, opts.max_chars);
  const more =
    text.length > opts.max_chars && doc.pages.length && !opts.pages
      ? `\n[Tip: jump with pages="…" or search with query="…" instead of reading everything.]`
      : "";
  return { content: [{ type: "text", text: `${head}${more}\n\n${body}` }] };
}

function tooLarge(size: number, limit: number) {
  return {
    content: [{ type: "text", text: `Error: File too large (${size} bytes, limit ${limit}).` }],
    isError: true,
  };
}

function notFoundOr(prefix: string, path: string, err: any) {
  const msg = err?.message ?? String(err);
  if (msg.includes("404") || msg.includes("Not Found")) {
    return { content: [{ type: "text", text: `Error: File not found: ${path}` }], isError: true };
  }
  return { content: [{ type: "text", text: `${prefix}: ${msg}` }], isError: true };
}

export function registerFileTools(server: McpServer, api: NextcloudAPI): void {
  // -----------------------------------------------------------------------
  // list_files
  // -----------------------------------------------------------------------
  server.registerTool(
    "list_files",
    {
      description: "List files and folders in a Nextcloud directory.",
      inputSchema: z.object({
        path: z
          .string()
          .optional()
          .default("/")
          .describe("Directory path to list (default: root /)"),
      }),
    },
    async ({ path }) => {
      try {
        const items = await api.listFiles(path ?? "/");
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(items, null, 2),
            },
          ],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: Directory not found: ${path}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error listing files: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // get_file
  // -----------------------------------------------------------------------
  server.registerTool(
    "get_file",
    {
      description:
        "Read the content of a text file from Nextcloud. Excel files (.xlsx/.xlsm) are converted to Markdown tables of all sheets (use read_spreadsheet for sheet/range selection). Word, PDF, OpenDocument (.odt/.ods/.odp) and PowerPoint (.pptx) files are converted to text (use read_document for page selection or search). ZIP archives such as EVEBI project files (.evex) list their entries (read them with read_archive). For other binary files, returns file metadata instead. Long texts are returned in windows (offset, max_chars).",
      inputSchema: z.object({
        path: z.string().describe("Path to the file"),
        ...pagingSchema,
      }),
    },
    async ({ path, offset, max_chars }) => {
      try {
        const info = await api.getFileInfo(path);
        if (info.is_directory) {
          return {
            content: [{ type: "text", text: `Error: '${path}' is a directory, not a file.` }],
            isError: true,
          };
        }
        if (isSpreadsheetFile(info.name)) {
          if (info.size > MAX_SPREADSHEET_BYTES) {
            return {
              content: [
                {
                  type: "text",
                  text: `Error: Spreadsheet too large (${info.size} bytes, limit ${MAX_SPREADSHEET_BYTES}).`,
                },
              ],
              isError: true,
            };
          }
          const buffer = await api.getFileBuffer(path);
          const text = await readSpreadsheet(buffer, info.name);
          return { content: [{ type: "text", text: page(text, offset, max_chars) }] };
        }
        if (
          isDocumentFile(info.name, info.content_type) ||
          /\.(odt|ods|odp|pptx)$/i.test(info.name)
        ) {
          if (info.size > MAX_DOCUMENT_BYTES) return tooLarge(info.size, MAX_DOCUMENT_BYTES);
          return await readExtracted(api, path, info, { offset, max_chars });
        }
        if (isArchiveFile(info.name)) {
          if (info.size > MAX_DOCUMENT_BYTES) return tooLarge(info.size, MAX_DOCUMENT_BYTES);
          const entries = await listArchive(await api.getFileBuffer(path));
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify(
                  {
                    message: "ZIP archive: read an entry with read_archive (path, entry).",
                    file_info: info,
                    entries,
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }
        if (!api.isTextFile(info.content_type, info.name)) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify(
                  {
                    message:
                      "File is binary or has an unsupported content type — cannot return as text.",
                    file_info: info,
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }
        const content = await api.getFileContent(path);
        return {
          content: [{ type: "text", text: page(content ?? "", offset, max_chars) }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: File not found: ${path}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error reading file: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // read_spreadsheet
  // -----------------------------------------------------------------------
  server.registerTool(
    "read_spreadsheet",
    {
      description:
        "Read the contents of an Excel file (.xlsx, .xlsm, .xltx, .xltm) from Nextcloud as text. " +
        "Returns calculated cell values (formula results, dates as ISO) as a Markdown table, CSV or JSON. " +
        "Optionally select a single sheet and/or an A1 range. Legacy .xls is not supported; read .ods with read_document.",
      inputSchema: z.object({
        path: z.string().describe("Path to the Excel file"),
        sheet: z
          .string()
          .optional()
          .describe("Sheet name or 1-based sheet index. Omit to read all sheets."),
        range: z
          .string()
          .optional()
          .describe('A1-style cell range, e.g. "A1:H100". Omit to read the used area.'),
        format: z
          .enum(["markdown", "csv", "json"])
          .optional()
          .default("markdown")
          .describe("Output format (default: markdown)"),
        max_rows: z
          .number()
          .int()
          .min(1)
          .max(10000)
          .optional()
          .default(500)
          .describe("Maximum data rows per sheet, excluding the header row (default: 500)"),
        include_formulas: z
          .boolean()
          .optional()
          .default(false)
          .describe("Append the formula to each calculated value, e.g. '42 [=SUM(A1:A3)]'"),
      }),
    },
    async ({ path, sheet, range, format, max_rows, include_formulas }) => {
      try {
        const info = await api.getFileInfo(path);
        if (info.is_directory) {
          return {
            content: [{ type: "text", text: `Error: '${path}' is a directory, not a file.` }],
            isError: true,
          };
        }
        if (isUnsupportedSpreadsheetFile(info.name)) {
          return {
            content: [
              {
                type: "text",
                text: `Error: '${info.name}' uses a legacy/unsupported spreadsheet format. Only .xlsx, .xlsm, .xltx and .xltm can be read here. For .ods use read_document (sheets as text).`,
              },
            ],
            isError: true,
          };
        }
        if (!isSpreadsheetFile(info.name)) {
          return {
            content: [
              {
                type: "text",
                text: `Error: '${info.name}' is not an Excel file (.xlsx/.xlsm/.xltx/.xltm). Use get_file for text files.`,
              },
            ],
            isError: true,
          };
        }
        if (info.size > MAX_SPREADSHEET_BYTES) {
          return {
            content: [
              {
                type: "text",
                text: `Error: Spreadsheet too large (${info.size} bytes, limit ${MAX_SPREADSHEET_BYTES}).`,
              },
            ],
            isError: true,
          };
        }
        const buffer = await api.getFileBuffer(path);
        const text = await readSpreadsheet(buffer, info.name, {
          sheet,
          range,
          format,
          maxRows: max_rows,
          includeFormulas: include_formulas,
        });
        return { content: [{ type: "text", text }] };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: File not found: ${path}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error reading spreadsheet: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // read_document
  // -----------------------------------------------------------------------
  server.registerTool(
    "read_document",
    {
      description:
        "Read the TEXT of a document in Nextcloud without loading the whole file into the context: PDF, Word " +
        "(.docx/.docm/.dotx), OpenDocument (.odt/.ods/.odp), PowerPoint (.pptx), Excel (.xlsx) and text files. " +
        "Three ways to read: (1) a window of the running text (offset/max_chars, default 15000 – the reply says " +
        'where to continue); (2) pages="3" or "2-5,8" for PDF pages, slides or sheets; (3) query="Frist Betrag" ' +
        "returns only the matching passages (±300 characters) with page and offset. The first line states the " +
        "number of pages and characters. Scanned PDFs without a text layer are flagged with a warning (needs " +
        "OCR, e.g. Paperless). The extracted text is cached, so further calls are fast.",
      inputSchema: z.object({
        path: z.string().describe("Path to the document"),
        pages: z
          .string()
          .optional()
          .describe('Pages/slides/sheets to read, e.g. "3", "1-5,8" or "10-" (default: all)'),
        query: z
          .string()
          .optional()
          .describe(
            "Search words (case-insensitive, any word matches); returns passages instead of the text"
          ),
        ...pagingSchema,
      }),
    },
    async ({ path, pages, query, offset, max_chars }) => {
      try {
        const info = await api.getFileInfo(path);
        if (info.is_directory) {
          return {
            content: [{ type: "text", text: `Error: '${path}' is a directory, not a file.` }],
            isError: true,
          };
        }
        if (!isExtractable(info.name, info.content_type)) {
          return {
            content: [
              {
                type: "text",
                text: `Error: '${info.name}' is not a supported document (PDF, Word, OpenDocument, PowerPoint, Excel, text). Images and scans need OCR (Paperless).`,
              },
            ],
            isError: true,
          };
        }
        if (info.size > MAX_DOCUMENT_BYTES) return tooLarge(info.size, MAX_DOCUMENT_BYTES);
        return await readExtracted(api, path, info, { pages, query, offset, max_chars });
      } catch (err: any) {
        return notFoundOr("Error reading document", path, err);
      }
    }
  );

  // -----------------------------------------------------------------------
  // read_archive
  // -----------------------------------------------------------------------
  server.registerTool(
    "read_archive",
    {
      description:
        "List or read entries of a ZIP archive in Nextcloud (.zip, EVEBI project files .evex/.evea, OpenDocument). " +
        "Without entry: list of entries with sizes. With entry: its text (XML and text files as is, Word/PDF " +
        "converted), e.g. entry 'projekt.xml' of an .evex. Long texts are returned in windows (offset, max_chars).",
      inputSchema: z.object({
        path: z.string().describe("Path to the archive"),
        entry: z
          .string()
          .optional()
          .describe("Entry to read, e.g. projekt.xml (omit to list entries)"),
        ...pagingSchema,
      }),
    },
    async ({ path, entry, offset, max_chars }) => {
      try {
        const info = await api.getFileInfo(path);
        if (info.is_directory) {
          return {
            content: [{ type: "text", text: `Error: '${path}' is a directory, not a file.` }],
            isError: true,
          };
        }
        if (info.size > MAX_DOCUMENT_BYTES) return tooLarge(info.size, MAX_DOCUMENT_BYTES);
        const buffer = await api.getFileBuffer(path);
        if (!entry) {
          const entries = await listArchive(buffer);
          return { content: [{ type: "text", text: JSON.stringify({ path, entries }, null, 2) }] };
        }
        const text = await readArchiveEntry(buffer, entry);
        return { content: [{ type: "text", text: page(text, offset, max_chars) }] };
      } catch (err: any) {
        return notFoundOr("Error reading archive", path, err);
      }
    }
  );

  // -----------------------------------------------------------------------
  // get_file_info
  // -----------------------------------------------------------------------
  server.registerTool(
    "get_file_info",
    {
      description:
        "Get metadata for a file or folder in Nextcloud (size, content type, modified date, etag, is_directory).",
      inputSchema: z.object({
        path: z.string().describe("Path to the file or folder"),
      }),
    },
    async ({ path }) => {
      try {
        const info = await api.getFileInfo(path);
        return {
          content: [{ type: "text", text: JSON.stringify(info, null, 2) }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: Path not found: ${path}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error getting file info: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // upload_file
  // -----------------------------------------------------------------------
  server.registerTool(
    "upload_file",
    {
      description: "Create or overwrite a text file in Nextcloud.",
      inputSchema: z.object({
        path: z.string().describe("Destination path for the file"),
        content: z.string().describe("Text content to write to the file"),
      }),
    },
    async ({ path, content }) => {
      try {
        await api.uploadFile(path, content);
        return {
          content: [{ type: "text", text: `File uploaded successfully: ${path}` }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        return {
          content: [{ type: "text", text: `Error uploading file: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // create_folder
  // -----------------------------------------------------------------------
  server.registerTool(
    "create_folder",
    {
      description: "Create a new folder (directory) in Nextcloud.",
      inputSchema: z.object({
        path: z.string().describe("Path of the folder to create"),
      }),
    },
    async ({ path }) => {
      try {
        await api.createFolder(path);
        return {
          content: [{ type: "text", text: `Folder created successfully: ${path}` }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("405") || msg.includes("Method Not Allowed")) {
          return {
            content: [{ type: "text", text: `Error: Folder already exists: ${path}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error creating folder: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // delete_file
  // -----------------------------------------------------------------------
  server.registerTool(
    "delete_file",
    {
      description: "Delete a file or folder from Nextcloud.",
      inputSchema: z.object({
        path: z.string().describe("Path to the file or folder to delete"),
      }),
    },
    async ({ path }) => {
      try {
        await api.deleteFile(path);
        return {
          content: [{ type: "text", text: `Deleted successfully: ${path}` }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: Path not found: ${path}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error deleting: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // move_file
  // -----------------------------------------------------------------------
  server.registerTool(
    "move_file",
    {
      description: "Move or rename a file or folder in Nextcloud.",
      inputSchema: z.object({
        from: z.string().describe("Source path"),
        to: z.string().describe("Destination path"),
      }),
    },
    async ({ from, to }) => {
      try {
        await api.moveFile(from, to);
        return {
          content: [{ type: "text", text: `Moved successfully: ${from} → ${to}` }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: Source not found: ${from}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error moving: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // copy_file
  // -----------------------------------------------------------------------
  server.registerTool(
    "copy_file",
    {
      description: "Copy a file or folder in Nextcloud.",
      inputSchema: z.object({
        from: z.string().describe("Source path"),
        to: z.string().describe("Destination path"),
      }),
    },
    async ({ from, to }) => {
      try {
        await api.copyFile(from, to);
        return {
          content: [{ type: "text", text: `Copied successfully: ${from} → ${to}` }],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        if (msg.includes("404") || msg.includes("Not Found")) {
          return {
            content: [{ type: "text", text: `Error: Source not found: ${from}` }],
            isError: true,
          };
        }
        return {
          content: [{ type: "text", text: `Error copying: ${msg}` }],
          isError: true,
        };
      }
    }
  );

  // -----------------------------------------------------------------------
  // search_files
  // -----------------------------------------------------------------------
  server.registerTool(
    "search_files",
    {
      description:
        "Search for files and folders by name in Nextcloud (case-insensitive filename match).",
      inputSchema: z.object({
        query: z.string().describe("Search string to match against filenames"),
        path: z
          .string()
          .optional()
          .default("/")
          .describe("Root path to search within (default: /)"),
      }),
    },
    async ({ query, path }) => {
      try {
        const results = await api.searchFiles(query, path ?? "/");
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { query, search_path: path ?? "/", count: results.length, results },
                null,
                2
              ),
            },
          ],
        };
      } catch (err: any) {
        const msg = err?.message ?? String(err);
        return {
          content: [{ type: "text", text: `Error searching files: ${msg}` }],
          isError: true,
        };
      }
    }
  );
}
