"use strict";
// Run with: npm test  (builds first, then node --test)
const test = require("node:test");
const assert = require("node:assert/strict");
const {
  extractDocument,
  cachedExtract,
  clearExtractCache,
  findSnippets,
  parsePageSpec,
  selectPages,
  isExtractable,
} = require("../build/utils/extract");
const { page, DEFAULT_MAX_CHARS } = require("../build/utils/documents");
const { makePdf, makeOdt, makeOds, makeOdp, makePptx, makeDocx } = require("./fixtures");

function longPdf(n = 60) {
  const pages = [];
  for (let p = 1; p <= n; p++) {
    const lines = Array.from({ length: 40 }, (_, i) => `Seite ${p} Zeile ${i}: Wärmedämmung U-Wert 0,${String(p).padStart(2, "0")}`);
    if (p === 42) lines[10] = "Förderbetrag gesamt: 12.345,67 EUR";
    pages.push(lines);
  }
  return makePdf(pages);
}

test("PDF: pages, umlauts, page offsets", async () => {
  const doc = await extractDocument(makePdf([["Größe: 5 m²", "Straße 1"], ["Zweite Seite"]]), "a.pdf");
  assert.equal(doc.kind, "pdf");
  assert.equal(doc.pages.length, 2);
  assert.ok(doc.text.includes("Größe: 5 m²"), doc.text);
  assert.ok(doc.text.includes("Straße 1"));
  assert.equal(doc.text.indexOf("--- Page 2 of 2 ---"), doc.pages[1].start);
  assert.equal(doc.needsOcr, false);
});

test("PDF: scans are flagged as needing OCR", async () => {
  const doc = await extractDocument(makePdf([null, null, ["Deckblatt mit etwas Text"]]), "scan.pdf");
  assert.equal(doc.needsOcr, true);
  assert.ok(doc.warnings.some((w) => w.includes("OCR")));
});

test("PDF: select pages and search with page numbers", async () => {
  const doc = await extractDocument(longPdf(), "bescheid.pdf");
  assert.equal(doc.pages.length, 60);
  const p42 = selectPages(doc, [42]);
  assert.ok(p42.startsWith("--- Page 42 of 60 ---"));
  assert.ok(!p42.includes("Page 43"));
  const two = selectPages(doc, [3, 7]);
  assert.ok(two.includes("--- Page 3 of 60 ---") && two.includes("--- Page 7 of 60 ---"));
  const { total, hits } = findSnippets(doc, "Förderbetrag");
  assert.equal(total, 1);
  assert.equal(hits[0].page, 42);
  assert.ok(hits[0].snippet.includes("12.345,67 EUR"));
});

test("windowing: default is small and the reply says how to continue", async () => {
  const doc = await extractDocument(longPdf(), "bescheid.pdf");
  assert.equal(DEFAULT_MAX_CHARS, 15000);
  const w = page(doc.text, 0, DEFAULT_MAX_CHARS);
  assert.ok(w.includes("Continue with offset=15000"));
  // stitching all windows gives back the full text
  let out = "";
  for (let off = 0; off < doc.text.length; off += 20000) out += doc.text.slice(off, off + 20000);
  assert.equal(out, doc.text);
});

test("ODT", async () => {
  const doc = await extractDocument(await makeOdt(["Löschkonzept", "Frist 10 Jahre"]), "konzept.odt");
  assert.equal(doc.text, "Löschkonzept\nFrist 10 Jahre");
});

test("ODS: sheets as pages, LibreOffice padding is not expanded", async () => {
  const buf = await makeOds({ "FLS 2025": [["Klient", "Stunden"], ["A", "12,5"]], Leer: [] });
  const doc = await extractDocument(buf, "fls.ods");
  assert.equal(doc.pages.length, 2);
  assert.ok(doc.text.includes("--- Sheet 1: FLS 2025 ---\nKlient\tStunden\nA\t12,5"), doc.text);
  assert.ok(doc.text.length < 200);
});

test("ODP and PPTX slides", async () => {
  const odp = await extractDocument(await makeOdp([["Titel"], ["Zweite Folie", "Punkt"]]), "v.odp");
  assert.equal(odp.pages.length, 2);
  assert.ok(odp.text.includes("--- Slide 2: Folie2 ---\nZweite Folie\nPunkt"), odp.text);
  const pptx = await extractDocument(await makePptx([["A"], ["B &amp; C"]]), "v.pptx");
  assert.ok(pptx.text.includes("--- Slide 2 ---\nB & C"), pptx.text);
});

test("DOCX via existing Word parser", async () => {
  const doc = await extractDocument(await makeDocx(["Angebot 4711"]), "a.docx");
  assert.equal(doc.text, "Angebot 4711");
});

test("page spec parsing", () => {
  assert.deepEqual(parsePageSpec("3", 10), [3]);
  assert.deepEqual(parsePageSpec("2-4,7", 10), [2, 3, 4, 7]);
  assert.deepEqual(parsePageSpec("8-", 10), [8, 9, 10]);
  assert.deepEqual(parsePageSpec("12", 10), []);
  assert.throws(() => parsePageSpec("x", 10));
});

test("supported types", () => {
  for (const n of ["a.pdf", "a.docx", "a.odt", "a.ods", "a.odp", "a.pptx", "a.xlsx", "a.txt"]) assert.ok(isExtractable(n), n);
  assert.ok(!isExtractable("foto.jpg"));
});

test("cache: second read does not load the file again", async () => {
  clearExtractCache();
  let loads = 0;
  const load = async () => {
    loads += 1;
    return { buffer: makePdf([["Eins"], ["Zwei"]]), filename: "x.pdf" };
  };
  const a = await cachedExtract("/x.pdf|etag1|100", load);
  const b = await cachedExtract("/x.pdf|etag1|100", load);
  assert.equal(loads, 1);
  assert.equal(a.cached, false);
  assert.equal(b.cached, true);
  await cachedExtract("/x.pdf|etag2|100", load); // changed file → new etag → reload
  assert.equal(loads, 2);
});
