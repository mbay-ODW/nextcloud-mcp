#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { randomUUID } from "node:crypto";
import express from "express";
import { parseArgs } from "node:util";
import { NextcloudAPI } from "./api/NextcloudAPI";
import { registerFileTools } from "./tools/files";
const { version } = require("../package.json") as { version: string };

// ---------------------------------------------------------------------------
// Logging – LOG_LEVEL env var: error|warn|info|debug|trace (default: info)
// ---------------------------------------------------------------------------
const LOG_LEVELS = ["error", "warn", "info", "debug", "trace"] as const;
type LogLevel = (typeof LOG_LEVELS)[number];
const requestedLevel = (process.env.LOG_LEVEL ?? "info").toLowerCase() as LogLevel;
const activeLevelIdx = Math.max(
  0,
  LOG_LEVELS.indexOf(LOG_LEVELS.includes(requestedLevel) ? requestedLevel : "info")
);
function lvlEnabled(l: LogLevel): boolean {
  return LOG_LEVELS.indexOf(l) <= activeLevelIdx;
}
const ts = () => new Date().toISOString();
const log = {
  error: (...a: unknown[]) =>
    lvlEnabled("error") && console.error(`[${ts()}] ERROR`, ...a),
  warn: (...a: unknown[]) =>
    lvlEnabled("warn") && console.warn(`[${ts()}] WARN `, ...a),
  info: (...a: unknown[]) =>
    lvlEnabled("info") && console.log(`[${ts()}] INFO `, ...a),
  debug: (...a: unknown[]) =>
    lvlEnabled("debug") && console.log(`[${ts()}] DEBUG`, ...a),
  trace: (...a: unknown[]) =>
    lvlEnabled("trace") && console.log(`[${ts()}] TRACE`, ...a),
};

const {
  values: { http: useHttp, port },
} = parseArgs({
  options: {
    http: { type: "boolean", default: false },
    port: { type: "string" },
  },
  allowPositionals: true,
});

const resolvedPort = port ? parseInt(port, 10) : 3000;

const nextcloudUrl = process.env.NEXTCLOUD_URL;
const nextcloudUsername = process.env.NEXTCLOUD_USERNAME;
const nextcloudAppPassword = process.env.NEXTCLOUD_APP_PASSWORD;

if (!nextcloudUrl || !nextcloudUsername || !nextcloudAppPassword) {
  console.error(
    "Missing required environment variables: NEXTCLOUD_URL, NEXTCLOUD_USERNAME, NEXTCLOUD_APP_PASSWORD"
  );
  process.exit(1);
}

async function main() {
  log.info(
    `nextcloud-mcp v${version} starting – LOG_LEVEL=${LOG_LEVELS[activeLevelIdx]} transport=${useHttp ? "http" : "stdio"} port=${resolvedPort}`
  );
  log.debug(
    `nextcloud: url=${nextcloudUrl} username=${nextcloudUsername}`
  );

  const api = new NextcloudAPI(nextcloudUrl!, nextcloudUsername!, nextcloudAppPassword!);

  const server = new McpServer(
    { name: "nextcloud-mcp", version },
    {
      instructions: `
Nextcloud MCP Server

Provides access to files and folders in a Nextcloud instance via WebDAV.

Available operations:
- list_files: List contents of a directory
- get_file: Read text file content
- get_file_info: Get file/folder metadata
- upload_file: Create or overwrite a text file
- create_folder: Create a new directory
- delete_file: Delete a file or folder
- move_file: Move or rename a file/folder
- copy_file: Copy a file/folder
- search_files: Search by filename (case-insensitive)

All paths are relative to the Nextcloud user's home directory root (/).
      `.trim(),
    }
  );

  registerFileTools(server, api);
  log.info("Tool group registered: files (9 tools)");

  if (useHttp) {
    const app = express();
    app.use(express.json());

    // Generic request log – sees EVERY incoming request before routing.
    app.use((req, _res, next) => {
      const auth = req.headers.authorization ?? "";
      const authPreview = auth
        ? auth.slice(0, 20) + (auth.length > 20 ? "…" : "")
        : "(none)";
      log.debug(
        `→ ${req.method} ${req.originalUrl}  auth=${authPreview}  ip=${req.ip}`
      );
      next();
    });

    // ------------------------------------------------------------------
    // OIDC / Bearer auth middleware
    //   1. If MCP_API_KEY is unset → all requests pass (dev mode, warned)
    //   2. Bearer == MCP_API_KEY → ok
    //   3. Bearer JWT → introspect against Authelia → active=true → ok
    //   4. Otherwise → 401
    // ------------------------------------------------------------------
    const mcpApiKey = process.env.MCP_API_KEY ?? "";
    const oidcIntrospectionUrl = process.env.OIDC_INTROSPECTION_URL ?? "";
    const oidcClientId = process.env.OIDC_CLIENT_ID ?? "";
    const oidcClientSecret = process.env.OIDC_CLIENT_SECRET ?? "";
    const oauthIssuer = process.env.OAUTH_ISSUER ?? "";
    const mcpServerUrl = process.env.MCP_SERVER_URL ?? "";

    log.info(
      `[auth] config: MCP_API_KEY=${mcpApiKey ? `set(${mcpApiKey.length} chars)` : "NOT SET"} OIDC_INTROSPECTION_URL=${oidcIntrospectionUrl || "NOT SET"} OIDC_CLIENT_ID=${oidcClientId || "NOT SET"} OIDC_CLIENT_SECRET=${oidcClientSecret ? `set(${oidcClientSecret.length} chars)` : "NOT SET"}`
    );
    log.info(
      `[oauth-discovery] OAUTH_ISSUER=${oauthIssuer || "NOT SET"} MCP_SERVER_URL=${mcpServerUrl || "NOT SET"}`
    );
    if (!mcpApiKey && (!oidcIntrospectionUrl || !oidcClientId || !oidcClientSecret)) {
      log.warn(
        "[auth] NEITHER static MCP_API_KEY NOR a complete OIDC triple is configured – ALL requests will be accepted unauthenticated (dev mode)."
      );
    }

    type AuthResult = { ok: true } | { ok: false; reason: "no_header" | "invalid_token" };

    const isAuthorized = async (req: express.Request): Promise<AuthResult> => {
      const tag = `${req.method} ${req.path}`;
      if (!mcpApiKey) {
        log.debug(`[auth] ${tag} – no MCP_API_KEY configured, passing through`);
        return { ok: true };
      }

      const auth = req.headers.authorization ?? "";
      const authPreview = auth
        ? auth.slice(0, 20) + (auth.length > 20 ? "…" : "")
        : "(none)";
      log.debug(`[auth] ${tag} – Authorization: ${authPreview}`);

      if (!auth) {
        log.warn(`[auth] ${tag} – DENY: no Authorization header`);
        return { ok: false, reason: "no_header" };
      }

      if (auth === `Bearer ${mcpApiKey}`) {
        log.info(`[auth] ${tag} – OK: static MCP_API_KEY matched`);
        return { ok: true };
      }

      if (!auth.startsWith("Bearer ")) {
        log.warn(`[auth] ${tag} – DENY: Authorization is not a Bearer scheme`);
        return { ok: false, reason: "invalid_token" };
      }

      // Bearer ≠ MCP_API_KEY → try OIDC introspection.
      if (!oidcIntrospectionUrl || !oidcClientId || !oidcClientSecret) {
        log.warn(
          `[auth] ${tag} – DENY: Bearer JWT presented but OIDC introspection not fully configured (url=${
            oidcIntrospectionUrl ? "ok" : "MISSING"
          } id=${oidcClientId ? "ok" : "MISSING"} secret=${
            oidcClientSecret ? "ok" : "MISSING"
          })`
        );
        return { ok: false, reason: "invalid_token" };
      }

      const jwtToken = auth.slice(7);
      log.debug(
        `[auth] ${tag} – introspecting token (len=${jwtToken.length}) against ${oidcIntrospectionUrl}`
      );
      const startedAt = Date.now();
      try {
        const credentials = Buffer.from(
          `${oidcClientId}:${oidcClientSecret}`
        ).toString("base64");
        const resp = await fetch(oidcIntrospectionUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
            Authorization: `Basic ${credentials}`,
          },
          body: `token=${encodeURIComponent(jwtToken)}`,
          signal: AbortSignal.timeout(5000),
        });
        const elapsed = Date.now() - startedAt;
        const body = await resp.text();
        log.debug(
          `[auth] ${tag} – introspection HTTP ${resp.status} in ${elapsed}ms, body: ${body.slice(0, 300)}`
        );
        if (resp.status !== 200) {
          log.warn(
            `[auth] ${tag} – DENY: introspection returned non-200 (${resp.status})`
          );
          return { ok: false, reason: "invalid_token" };
        }
        const data = JSON.parse(body) as {
          active?: boolean;
          sub?: string;
          scope?: string;
          aud?: string;
          exp?: number;
        };
        if (data.active === true) {
          log.info(
            `[auth] ${tag} – OK: OIDC token active sub=${data.sub ?? "?"} scope=${data.scope ?? "?"}`
          );
          return { ok: true };
        }
        log.warn(`[auth] ${tag} – DENY: OIDC token not active`);
        return { ok: false, reason: "invalid_token" };
      } catch (e) {
        log.error(`[auth] ${tag} – introspection exception:`, e);
        return { ok: false, reason: "invalid_token" };
      }
    };

    /**
     * Send a 401.
     *   - reason "invalid_token" → emit an RFC 6750 WWW-Authenticate
     *     `Bearer error="invalid_token"` challenge. This triggers
     *     Claude.ai's silent refresh-token flow.
     *   - reason "no_header" → emit a "naked" 401 WITHOUT a
     *     WWW-Authenticate header. Sending `Bearer realm="…"` here
     *     would short-circuit Claude.ai's OAuth discovery.
     */
    const sendUnauthorized = (
      res: express.Response,
      reason: "no_header" | "invalid_token" = "no_header"
    ) => {
      if (reason === "invalid_token") {
        res.set(
          "WWW-Authenticate",
          'Bearer realm="nextcloud-mcp", error="invalid_token", error_description="The access token expired or is invalid"'
        );
      }
      res.status(401).json({ error: "Unauthorized" });
    };

    const authMiddleware: express.RequestHandler = async (req, res, next) => {
      const result = await isAuthorized(req);
      if (result.ok) return next();
      sendUnauthorized(res, result.reason);
    };

    const sseTransports: Record<string, SSEServerTransport> = {};

    // ------------------------------------------------------------------
    // Public OAuth discovery endpoints (RFC 9728 + RFC 8414).
    // Mounted BEFORE auth so Claude.ai can bootstrap the OAuth flow.
    // ------------------------------------------------------------------
    if (oauthIssuer && mcpServerUrl) {
      app.get("/.well-known/oauth-protected-resource", (_req, res) => {
        log.info("[discovery] /.well-known/oauth-protected-resource hit");
        res.json({
          resource: mcpServerUrl,
          authorization_servers: [oauthIssuer],
          bearer_methods_supported: ["header"],
          scopes_supported: ["openid", "profile", "email"],
        });
      });
    }
    if (oauthIssuer) {
      app.get("/.well-known/oauth-authorization-server", async (_req, res) => {
        log.info("[discovery] /.well-known/oauth-authorization-server hit");
        try {
          const upstream = await fetch(
            `${oauthIssuer}/.well-known/oauth-authorization-server`,
            { signal: AbortSignal.timeout(5000) }
          );
          if (upstream.ok) {
            const data = await upstream.json();
            return res.json(data);
          }
        } catch (e) {
          log.warn("[discovery] upstream oauth-authorization-server fetch failed, using built-in response:", e);
        }
        // Fallback: build from known Authelia endpoints
        res.json({
          issuer: oauthIssuer,
          authorization_endpoint: `${oauthIssuer}/api/oidc/authorization`,
          token_endpoint: `${oauthIssuer}/api/oidc/token`,
          jwks_uri: `${oauthIssuer}/jwks.json`,
          introspection_endpoint: `${oauthIssuer}/api/oidc/introspection`,
          response_types_supported: ["code"],
          grant_types_supported: ["authorization_code", "refresh_token"],
          code_challenge_methods_supported: ["S256"],
          scopes_supported: ["openid", "profile", "email"],
        });
      });
    }

    // ------------------------------------------------------------------
    // Streamable HTTP transport (current MCP spec) – STATEFUL.
    // Mounted on BOTH /mcp and /sse (POST) for compatibility.
    // ------------------------------------------------------------------
    const streamableTransports: Record<string, StreamableHTTPServerTransport> = {};

    const streamableHandler: express.RequestHandler = async (req, res) => {
      const tag = `[stream POST ${req.path}]`;
      const incomingSessionId = (req.headers["mcp-session-id"] as
        | string
        | undefined) ?? undefined;
      log.debug(
        `${tag} session=${incomingSessionId ?? "(new)"} body keys=${Object.keys(req.body ?? {}).join(",")}`
      );
      log.trace(`${tag} body: ${JSON.stringify(req.body).slice(0, 500)}`);

      try {
        let transport: StreamableHTTPServerTransport;

        if (incomingSessionId && streamableTransports[incomingSessionId]) {
          transport = streamableTransports[incomingSessionId];
        } else if (!incomingSessionId && isInitializeRequest(req.body)) {
          transport = new StreamableHTTPServerTransport({
            sessionIdGenerator: () => randomUUID(),
            onsessioninitialized: (sid) => {
              log.info(`${tag} session initialized: ${sid}`);
              streamableTransports[sid] = transport;
            },
          });
          transport.onclose = () => {
            if (transport.sessionId) {
              log.info(
                `${tag} session closed: ${transport.sessionId} (remaining: ${
                  Object.keys(streamableTransports).length - 1
                })`
              );
              delete streamableTransports[transport.sessionId];
            }
          };
          await server.connect(transport);
        } else {
          log.warn(
            `${tag} bad request – no session id and not an initialize call`
          );
          res.status(400).json({
            jsonrpc: "2.0",
            error: {
              code: -32000,
              message: "Bad Request: no valid session id and not an initialize call",
            },
            id: null,
          });
          return;
        }

        await transport.handleRequest(req, res, req.body);
        log.debug(`${tag} handled, response status=${res.statusCode}`);
      } catch (error) {
        log.error(`${tag} handler error:`, error);
        if (!res.headersSent) {
          res.status(500).json({
            jsonrpc: "2.0",
            error: { code: -32603, message: "Internal server error" },
            id: null,
          });
        }
      }
    };

    const streamableSessionLookupHandler: express.RequestHandler = async (
      req,
      res
    ) => {
      const sid = req.headers["mcp-session-id"] as string | undefined;
      if (!sid || !streamableTransports[sid]) {
        res.status(400).json({
          jsonrpc: "2.0",
          error: { code: -32000, message: "Invalid or missing session id" },
          id: null,
        });
        return;
      }
      await streamableTransports[sid].handleRequest(req, res);
    };

    app.post("/mcp", authMiddleware, streamableHandler);
    app.post("/sse", authMiddleware, streamableHandler);
    app.get("/mcp", authMiddleware, streamableSessionLookupHandler);
    app.delete("/mcp", authMiddleware, streamableSessionLookupHandler);

    // ------------------------------------------------------------------
    // SSE transport (legacy MCP spec; Claude Desktop, older Claude.ai).
    // ------------------------------------------------------------------
    app.get("/sse", authMiddleware, async (_req, res) => {
      log.info("[/sse GET] new SSE connection");
      try {
        const transport = new SSEServerTransport("/messages", res);
        sseTransports[transport.sessionId] = transport;
        log.info(
          `[/sse GET] sessionId=${transport.sessionId} – transport registered (active sessions: ${Object.keys(sseTransports).length})`
        );
        res.on("close", () => {
          log.info(
            `[/sse GET] sessionId=${transport.sessionId} – connection closed (remaining sessions: ${
              Object.keys(sseTransports).length - 1
            })`
          );
          delete sseTransports[transport.sessionId];
          transport.close();
        });
        await server.connect(transport);
        log.debug(`[/sse GET] sessionId=${transport.sessionId} – server.connect completed`);
      } catch (error) {
        log.error("[/sse GET] error:", error);
        if (!res.headersSent) {
          res.status(500).json({
            jsonrpc: "2.0",
            error: { code: -32603, message: "Internal server error" },
            id: null,
          });
        }
      }
    });

    app.post("/messages", authMiddleware, async (req, res) => {
      const sessionId = req.query.sessionId as string;
      log.debug(
        `[/messages POST] sessionId=${sessionId} body keys=${Object.keys(req.body ?? {}).join(",")}`
      );
      log.trace(`[/messages POST] body: ${JSON.stringify(req.body).slice(0, 500)}`);
      const transport = sseTransports[sessionId];
      if (!transport) {
        log.warn(
          `[/messages POST] sessionId=${sessionId} – no transport found (known: ${Object.keys(sseTransports).join(",") || "none"})`
        );
        res.status(400).send("No transport found for sessionId");
        return;
      }
      try {
        await transport.handlePostMessage(req, res, req.body);
        log.debug(`[/messages POST] sessionId=${sessionId} – handled, status=${res.statusCode}`);
      } catch (error) {
        log.error(`[/messages POST] sessionId=${sessionId} – error:`, error);
        if (!res.headersSent) res.status(500).send("Internal server error");
      }
    });

    // Catch-all for unmatched routes
    app.use((req, res) => {
      log.warn(
        `[404] ${req.method} ${req.originalUrl} – no route matched. Headers: ${JSON.stringify(req.headers).slice(0, 300)}`
      );
      res.status(404).json({ error: "Not found", path: req.originalUrl });
    });

    app.listen(resolvedPort, () => {
      log.info(
        `MCP server listening on :${resolvedPort} (routes: /mcp /sse /messages /.well-known/*)`
      );
    });
  } else {
    // stdio mode (for local/Claude Desktop use)
    const { StdioServerTransport } = await import(
      "@modelcontextprotocol/sdk/server/stdio.js"
    );
    log.info("starting on stdio");
    const transport = new StdioServerTransport();
    await server.connect(transport);
  }
}

main().catch((e) => {
  log.error("fatal:", e);
  process.exit(1);
});
