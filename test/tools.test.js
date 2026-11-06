"use strict";
// End-to-end through the registered tools with a fake Nextcloud API.
const test = require("node:test");
const assert = require("node:assert/strict");
const { registerFileTools } = require("../build/tools/files");
const { clearExtractCache } = require("../build/utils/extract");
const { makePdf, makeOds } = require("./fixtures");

function setup(files) {
  const tools = {};
  const server = { registerTool: (name, _config, handler) => (tools[name] = handler) };
  const downloads = {};
  const api = {
    async getFileInfo(path) {
      const f = files[path];
      if (!f) throw new Error("404 Not Found");
      return { name: path.split("/").pop(), path, size: f.length, content_type: "application/octet-stream", etag: "e1", is_directory: false };
    },
    async getFileBuffer(path) {
      downloads[path] = (downloads[path] ?? 0) + 1;
      return files[path];
    },
    isTextFile: () => false,
  };
  registerFileTools(server, api);
  return { tools, downloads };
}

const call = (h, args) => h({ offset: 0, max_chars: 15000, ...args });

function bigPdf() {
  const pages = [];
  for (let p = 1; p <= 50; p++) {
    const lines = Array.from({ length: 40 }, (_, i) => `Seite ${p} Zeile ${i} Text Text Text`);
    if (p === 33) lines[5] = "Widerrufsfrist endet am 11.11.2026";
    pages.push(lines);
  }
  return makePdf(pages);
}

test("read_document: header, small window, pages, query, cache", async () => {
  clearExtractCache();
  const { tools, downloads } = setup({ "/Projekte/Bescheid.pdf": bigPdf() });
  const read = tools.read_document;

  const first = await call(read, { path: "/Projekte/Bescheid.pdf" });
  const text = first.content[0].text;
  assert.ok(text.startsWith("[Bescheid.pdf: PDF, 50 pages,"), text.slice(0, 120));
  assert.ok(text.includes('pages="…"'));
  assert.ok(text.includes("Continue with offset=15000"));
  assert.ok(text.length < 16000);

  const p33 = (await call(read, { path: "/Projekte/Bescheid.pdf", pages: "33" })).content[0].text;
  assert.ok(p33.includes("--- Page 33 of 50 ---") && !p33.includes("Page 34"));

  const q = JSON.parse((await call(read, { path: "/Projekte/Bescheid.pdf", query: "Widerrufsfrist" })).content[0].text);
  assert.equal(q.total_hits, 1);
  assert.equal(q.hits[0].page, 33);

  assert.equal(downloads["/Projekte/Bescheid.pdf"], 1, "file downloaded once, then served from cache");

  const bad = await call(read, { path: "/Projekte/Bescheid.pdf", pages: "99" });
  assert.equal(bad.isError, true);
});

test("read_document and get_file read .ods as text", async () => {
  clearExtractCache();
  const { tools } = setup({ "/FLS.ods": await makeOds({ Jahr: [["Klient", "FLS"], ["A", "40"]] }) });
  const r = (await call(tools.read_document, { path: "/FLS.ods" })).content[0].text;
  assert.ok(r.includes("Klient\tFLS\nA\t40"), r);
  const g = (await call(tools.get_file, { path: "/FLS.ods" })).content[0].text;
  assert.ok(g.includes("Klient\tFLS"), g);
});

test("read_document rejects images with a hint", async () => {
  const { tools } = setup({ "/foto.jpg": Buffer.from([0xff, 0xd8, 0xff]) });
  const r = await call(tools.read_document, { path: "/foto.jpg" });
  assert.equal(r.isError, true);
  assert.ok(r.content[0].text.includes("OCR"));
});
