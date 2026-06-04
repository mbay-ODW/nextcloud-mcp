// @ts-nocheck
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { NextcloudAPI } from "../api/NextcloudAPI";

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
    "Read the content of a text file from Nextcloud. For binary files, returns file metadata instead.",
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
