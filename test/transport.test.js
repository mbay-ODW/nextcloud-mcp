"use strict";
// Runs the built server over HTTP: both protocol eras, sessions and auth.
const test = require("node:test");
const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");
const path = require("node:path");

const TOOL_COUNT = 12;
const PORT = 41000 + (process.pid % 5000);
const base = `http://127.0.0.1:${PORT}`;
let child;

test.before(async () => {
  child = spawn(
    process.execPath,
    [path.join(__dirname, "../build/index.js"), "--http", "--port", String(PORT)],
    {
      env: {
        ...process.env,
        MCP_API_KEY: "testkey",
        NEXTCLOUD_URL: "http://127.0.0.1:9",
        NEXTCLOUD_USERNAME: "user",
        NEXTCLOUD_APP_PASSWORD: "password",
      },
      stdio: "ignore",
    }
  );
  for (let i = 0; i < 100; i++) {
    try {
      await fetch(`${base}/mcp`);
      return;
    } catch {
      await new Promise((r) => setTimeout(r, 100));
    }
  }
  throw new Error("server did not start");
});

test.after(() => child.kill());

const ENVELOPE = {
  "io.modelcontextprotocol/protocolVersion": "2026-07-28",
  "io.modelcontextprotocol/clientInfo": { name: "test", version: "1" },
  "io.modelcontextprotocol/clientCapabilities": {},
};
const INIT = {
  jsonrpc: "2.0",
  id: 1,
  method: "initialize",
  params: {
    protocolVersion: "2025-11-25",
    capabilities: {},
    clientInfo: { name: "test", version: "1" },
  },
};
const LIST = { jsonrpc: "2.0", id: 2, method: "tools/list" };
const modern = (method, params = {}) => ({
  jsonrpc: "2.0",
  id: 3,
  method,
  params: { ...params, _meta: ENVELOPE },
});

async function send(path, method, body, headers = {}, token = "testkey") {
  const res = await fetch(`${base}${path}`, {
    method,
    headers: {
      Accept: "application/json, text/event-stream",
      ...(body ? { "Content-Type": "application/json" } : {}),
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...headers,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  const sse = text.split("\n").filter((l) => l.startsWith("data:") && l.length > 6);
  const payload = sse.length ? sse[sse.length - 1].slice(5) : text;
  let json;
  try {
    json = JSON.parse(payload);
  } catch {
    json = undefined;
  }
  return { status: res.status, headers: res.headers, json };
}

test("2025 era: sessionful initialize, tools/list and session errors", async () => {
  for (const path of ["/mcp", "/sse"]) {
    const init = await send(path, "POST", INIT);
    assert.equal(init.status, 200);
    assert.equal(init.json.result.protocolVersion, "2025-11-25");
    const sid = init.headers.get("mcp-session-id");
    assert.ok(sid);

    const list = await send(path, "POST", LIST, { "Mcp-Session-Id": sid });
    assert.equal(list.json.result.tools.length, TOOL_COUNT);
    for (const tool of list.json.result.tools) {
      assert.ok(tool.description, `${tool.name} has a description`);
      assert.equal(tool.inputSchema.type, "object", `${tool.name} has an input schema`);
    }

    const gone = await send(path, "POST", LIST, { "Mcp-Session-Id": "unknown" });
    assert.equal(gone.status, 404);
    assert.equal((await send(path, "POST", LIST)).status, 400);

    const missing = await send(path, "GET");
    assert.equal(missing.status, 400);
    assert.equal(missing.json.error.message, "Missing session id");

    assert.equal((await send(path, "DELETE", undefined, { "Mcp-Session-Id": sid })).status, 200);
    assert.equal((await send(path, "POST", LIST, { "Mcp-Session-Id": sid })).status, 404);
  }
});

test("2025 era: initialize with a stale session id starts a new session", async () => {
  const init = await send("/mcp", "POST", INIT, { "Mcp-Session-Id": "stale" });
  assert.equal(init.status, 200);
  assert.notEqual(init.headers.get("mcp-session-id"), "stale");
});

test("2026-07-28: answered per request on /mcp and /sse, without a session", async () => {
  const headers = { "MCP-Protocol-Version": "2026-07-28" };
  for (const path of ["/mcp", "/sse"]) {
    const discover = await send(path, "POST", modern("server/discover"), {
      ...headers,
      "Mcp-Method": "server/discover",
    });
    assert.equal(discover.status, 200);
    assert.ok(discover.json.result.supportedVersions.includes("2026-07-28"));
    assert.equal(discover.headers.get("mcp-session-id"), null);

    const list = await send(path, "POST", modern("tools/list"), {
      ...headers,
      "Mcp-Method": "tools/list",
    });
    assert.equal(list.json.result.tools.length, TOOL_COUNT);
  }
});

test("auth applies to both eras and keeps its WWW-Authenticate semantics", async () => {
  for (const body of [INIT, modern("tools/list")]) {
    const none = await send("/mcp", "POST", body, {}, null);
    assert.equal(none.status, 401);
    assert.equal(none.headers.get("www-authenticate"), null);

    const wrong = await send("/mcp", "POST", body, {}, "wrong");
    assert.equal(wrong.status, 401);
    assert.match(wrong.headers.get("www-authenticate"), /error="invalid_token"/);
  }
});

test("a body above the SDK's 4 MiB default stays on the session path", async () => {
  const init = await send("/mcp", "POST", INIT);
  const sid = init.headers.get("mcp-session-id");
  const upload = await send(
    "/mcp",
    "POST",
    {
      jsonrpc: "2.0",
      id: 4,
      method: "tools/call",
      params: { name: "upload_file", arguments: { path: "/big.txt", content: "x".repeat(5 * 1024 * 1024) } },
    },
    { "Mcp-Session-Id": sid }
  );
  assert.equal(upload.status, 200);
  assert.equal(upload.json.id, 4);
  assert.ok(upload.json.result, "tool ran; the fake backend makes it fail in-band");
});
