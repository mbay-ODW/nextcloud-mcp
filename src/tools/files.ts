// @ts-nocheck
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { NextcloudAPI } from "../api/NextcloudAPI";
import {
  isSpreadsheetFile,
  isUnsupportedSpreadsheetFile,
  readSpreadsheet,
} from "../utils/spreadsheet";

const MAX_SPREADSHEET_BYTES = 50 * 1024 * 1024;

export function registerFileTools(server: McpServer, api: NextcloudAPI): void {
  // -----------------------------------------------------------------------
  // list_files
  // -----------------------------------------------------------------------
  server.tool(
    "list_files",
    "List files and folders in a Nextcloud directory.",
    {
      path: z
        .string()
        .optional()
        .default("/")
        .describe("Directory path to list (default: root /)"),
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
  server.tool(
    "get_file",
    "Read the content of a text file from Nextcloud. Excel files (.xlsx/.xlsm) are converted to Markdown tables of all sheets (use read_spreadsheet for sheet/range selection). For other binary files, returns file metadata instead.",
    {
      path: z.string().describe("Path to the file"),
    },
    async ({ path }) => {
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
              content: [{ type: "text", text: `Error: Spreadsheet too large (${info.size} bytes, limit ${MAX_SPREADSHEET_BYTES}).` }],
              isError: true,
            };
          }
          const buffer = await api.getFileBuffer(path);
          const text = await readSpreadsheet(buffer, info.name);
          return { content: [{ type: "text", text }] };
        }
        if (!api.isTextFile(info.content_type, info.name)) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify(
                  {
                    message: "File is binary or has an unsupported content type — cannot return as text.",
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
          content: [{ type: "text", text: content ?? "" }],
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
  server.tool(
    "read_spreadsheet",
    "Read the contents of an Excel file (.xlsx, .xlsm, .xltx, .xltm) from Nextcloud as text. " +
      "Returns calculated cell values (formula results, dates as ISO) as a Markdown table, CSV or JSON. " +
      "Optionally select a single sheet and/or an A1 range. Legacy .xls and .ods are not supported — save them as .xlsx first.",
    {
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
                text: `Error: '${info.name}' uses a legacy/unsupported spreadsheet format. Only .xlsx, .xlsm, .xltx and .xltm can be read — please save the file as .xlsx.`,
              },
            ],
            isError: true,
          };
        }
        if (!isSpreadsheetFile(info.name)) {
          return {
            content: [{ type: "text", text: `Error: '${info.name}' is not an Excel file (.xlsx/.xlsm/.xltx/.xltm). Use get_file for text files.` }],
            isError: true,
          };
        }
        if (info.size > MAX_SPREADSHEET_BYTES) {
          return {
            content: [{ type: "text", text: `Error: Spreadsheet too large (${info.size} bytes, limit ${MAX_SPREADSHEET_BYTES}).` }],
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
  // get_file_info
  // -----------------------------------------------------------------------
  server.tool(
    "get_file_info",
    "Get metadata for a file or folder in Nextcloud (size, content type, modified date, etag, is_directory).",
    {
      path: z.string().describe("Path to the file or folder"),
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
  server.tool(
    "upload_file",
    "Create or overwrite a text file in Nextcloud.",
    {
      path: z.string().describe("Destination path for the file"),
      content: z.string().describe("Text content to write to the file"),
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
  server.tool(
    "create_folder",
    "Create a new folder (directory) in Nextcloud.",
    {
      path: z.string().describe("Path of the folder to create"),
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
  server.tool(
    "delete_file",
    "Delete a file or folder from Nextcloud.",
    {
      path: z.string().describe("Path to the file or folder to delete"),
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
  server.tool(
    "move_file",
    "Move or rename a file or folder in Nextcloud.",
    {
      from: z.string().describe("Source path"),
      to: z.string().describe("Destination path"),
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
  server.tool(
    "copy_file",
    "Copy a file or folder in Nextcloud.",
    {
      from: z.string().describe("Source path"),
      to: z.string().describe("Destination path"),
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
  server.tool(
    "search_files",
    "Search for files and folders by name in Nextcloud (case-insensitive filename match).",
    {
      query: z.string().describe("Search string to match against filenames"),
      path: z
        .string()
        .optional()
        .default("/")
        .describe("Root path to search within (default: /)"),
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
