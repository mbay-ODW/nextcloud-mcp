import { createClient, WebDAVClient, FileStat, ResponseDataDetailed } from "webdav";

export interface FileItem {
  name: string;
  type: "file" | "directory";
  size: number;
  modified: string;
  path: string;
  etag?: string;
  content_type?: string;
}

export interface FileInfo {
  name: string;
  path: string;
  size: number;
  content_type: string;
  modified: string;
  etag: string;
  is_directory: boolean;
}

export class NextcloudAPI {
  private client: WebDAVClient;
  private baseUrl: string;
  private username: string;

  constructor(nextcloudUrl: string, username: string, appPassword: string) {
    this.baseUrl = nextcloudUrl.replace(/\/$/, "");
    this.username = username;
    const davUrl = `${this.baseUrl}/remote.php/dav/files/${encodeURIComponent(username)}`;
    this.client = createClient(davUrl, {
      username,
      password: appPassword,
    });
  }

  private normalizePath(p: string): string {
    // Ensure path starts with /
    if (!p.startsWith("/")) p = "/" + p;
    return p;
  }

  async listFiles(path: string = "/"): Promise<FileItem[]> {
    path = this.normalizePath(path);
    const contents = await this.client.getDirectoryContents(path, {
      deep: false,
    });
    const items = Array.isArray(contents)
      ? (contents as FileStat[])
      : (contents as ResponseDataDetailed<FileStat[]>).data;

    // Filter out the directory itself (first item when listing a dir)
    return items
      .filter((item) => {
        // Normalize paths for comparison
        const itemPath = item.filename.replace(/\/$/, "");
        const queryPath = path.replace(/\/$/, "");
        return itemPath !== queryPath;
      })
      .map((item) => ({
        name: item.basename,
        type: item.type === "directory" ? "directory" : "file",
        size: item.size ?? 0,
        modified: item.lastmod ?? "",
        path: item.filename,
        etag: item.etag ?? undefined,
        content_type: item.mime ?? undefined,
      }));
  }

  async getFileContent(path: string): Promise<string | null> {
    path = this.normalizePath(path);
    const content = await this.client.getFileContents(path, {
      format: "text",
    });
    if (typeof content === "string") return content;
    // ResponseDataDetailed
    const data = (content as ResponseDataDetailed<string>).data;
    return data;
  }

  async getFileInfo(path: string): Promise<FileInfo> {
    path = this.normalizePath(path);
    const stat = await this.client.stat(path) as FileStat | ResponseDataDetailed<FileStat>;
    const item: FileStat = "data" in stat ? (stat as ResponseDataDetailed<FileStat>).data : (stat as FileStat);
    return {
      name: item.basename,
      path: item.filename,
      size: item.size ?? 0,
      content_type: item.mime ?? "application/octet-stream",
      modified: item.lastmod ?? "",
      etag: item.etag ?? "",
      is_directory: item.type === "directory",
    };
  }

  async uploadFile(path: string, content: string): Promise<void> {
    path = this.normalizePath(path);
    await this.client.putFileContents(path, content, {
      overwrite: true,
    });
  }

  async createFolder(path: string): Promise<void> {
    path = this.normalizePath(path);
    await this.client.createDirectory(path, { recursive: false });
  }

  async deleteFile(path: string): Promise<void> {
    path = this.normalizePath(path);
    await this.client.deleteFile(path);
  }

  async moveFile(from: string, to: string): Promise<void> {
    from = this.normalizePath(from);
    to = this.normalizePath(to);
    await this.client.moveFile(from, to);
  }

  async copyFile(from: string, to: string): Promise<void> {
    from = this.normalizePath(from);
    to = this.normalizePath(to);
    await this.client.copyFile(from, to);
  }

  async searchFiles(query: string, searchPath: string = "/"): Promise<FileItem[]> {
    searchPath = this.normalizePath(searchPath);
    // Use getDirectoryContents with deep:true then filter by filename
    const contents = await this.client.getDirectoryContents(searchPath, {
      deep: true,
    });
    const items = Array.isArray(contents)
      ? (contents as FileStat[])
      : (contents as ResponseDataDetailed<FileStat[]>).data;

    const lowerQuery = query.toLowerCase();
    return items
      .filter((item) => item.basename.toLowerCase().includes(lowerQuery))
      .map((item) => ({
        name: item.basename,
        type: item.type === "directory" ? "directory" : "file",
        size: item.size ?? 0,
        modified: item.lastmod ?? "",
        path: item.filename,
        etag: item.etag ?? undefined,
        content_type: item.mime ?? undefined,
      }));
  }

  isTextFile(contentType: string | undefined, filename: string): boolean {
    if (contentType) {
      if (
        contentType.startsWith("text/") ||
        contentType === "application/json" ||
        contentType === "application/xml" ||
        contentType === "application/javascript" ||
        contentType === "application/typescript" ||
        contentType === "application/x-yaml" ||
        contentType === "application/yaml"
      ) {
        return true;
      }
    }
    // Fallback: check extension
    const textExtensions = [
      ".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm",
      ".js", ".ts", ".jsx", ".tsx", ".css", ".scss", ".sass",
      ".yaml", ".yml", ".toml", ".ini", ".conf", ".cfg",
      ".sh", ".bash", ".zsh", ".fish", ".py", ".rb", ".php",
      ".java", ".go", ".rs", ".c", ".cpp", ".h", ".hpp",
      ".sql", ".graphql", ".gql", ".env", ".log", ".diff",
      ".patch", ".rst", ".tex", ".svg",
    ];
    const ext = "." + filename.split(".").pop()?.toLowerCase();
    return textExtensions.includes(ext);
  }
}
