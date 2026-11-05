"use strict";
// Small, dependency-free builders for test documents (PDF via raw syntax, Office/ODF via JSZip).
const JSZip = require("jszip");

function cp1252(str) {
  // Latin-1 range covers the umlauts and ß used in the tests
  return Buffer.from(str, "latin1");
}

function pdfEscape(line) {
  const raw = cp1252(line).toString("latin1");
  return raw.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

/** PDF with one page per entry (array of lines); null = page without text (like a scan). */
function makePdf(pages) {
  const objects = [];
  const add = (s) => objects.push(s) && objects.length;
  const font = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
  const pagesId = objects.length + 1 + 2 * pages.length;
  const pageIds = [];
  for (const lines of pages) {
    let stream;
    if (lines && lines.length) {
      stream = ["BT /F1 11 Tf 14 TL 50 800 Td", ...lines.map((l) => `(${pdfEscape(l)}) Tj T*`), "ET"].join("\n");
    } else {
      stream = "0.9 g 50 50 400 700 re f";
    }
    const content = add(`<< /Length ${Buffer.byteLength(stream, "latin1")} >>\nstream\n${stream}\nendstream`);
    pageIds.push(
      add(
        `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${font} 0 R >> >> /Contents ${content} 0 R >>`
      )
    );
  }
  const kids = pageIds.map((i) => `${i} 0 R`).join(" ");
  if (add(`<< /Type /Pages /Kids [${kids}] /Count ${pageIds.length} >>`) !== pagesId) throw new Error("id mismatch");
  const catalog = add(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

  let out = "%PDF-1.4\n";
  const offsets = [];
  objects.forEach((obj, i) => {
    offsets.push(Buffer.byteLength(out, "latin1"));
    out += `${i + 1} 0 obj\n${obj}\nendobj\n`;
  });
  const xref = Buffer.byteLength(out, "latin1");
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const off of offsets) out += `${String(off).padStart(10, "0")} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length + 1} /Root ${catalog} 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}

async function zip(files, mimetype) {
  const z = new JSZip();
  if (mimetype) z.file("mimetype", mimetype, { compression: "STORE" });
  for (const [name, content] of Object.entries(files)) z.file(name, content);
  return z.generateAsync({ type: "nodebuffer" });
}

const ODF_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"';

function odfDoc(body) {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NS}><office:body>${body}</office:body></office:document-content>`;
}

function makeOdt(paragraphs) {
  const body = paragraphs.map((p) => `<text:p>${p}</text:p>`).join("");
  return zip({ "content.xml": odfDoc(`<office:text>${body}</office:text>`) }, "application/vnd.oasis.opendocument.text");
}

function makeOds(sheets) {
  let tables = "";
  for (const [name, rows] of Object.entries(sheets)) {
    let xml = rows
      .map(
        (row) =>
          "<table:table-row>" +
          row.map((c) => `<table:table-cell><text:p>${c}</text:p></table:table-cell>`).join("") +
          '<table:table-cell table:number-columns-repeated="16378"/></table:table-row>'
      )
      .join("");
    xml +=
      '<table:table-row table:number-rows-repeated="1048000"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>';
    tables += `<table:table table:name="${name}">${xml}</table:table>`;
  }
  return zip({ "content.xml": odfDoc(`<office:spreadsheet>${tables}</office:spreadsheet>`) }, "application/vnd.oasis.opendocument.spreadsheet");
}

function makeOdp(slides) {
  const pages = slides
    .map((s, i) => `<draw:page draw:name="Folie${i + 1}"><draw:frame><draw:text-box>${s.map((t) => `<text:p>${t}</text:p>`).join("")}</draw:text-box></draw:frame></draw:page>`)
    .join("");
  return zip({ "content.xml": odfDoc(`<office:presentation>${pages}</office:presentation>`) }, "application/vnd.oasis.opendocument.presentation");
}

function makePptx(slides) {
  const files = { "ppt/presentation.xml": "<p/>" };
  slides.forEach((paras, i) => {
    const ps = paras.map((t) => `<a:p><a:r><a:t>${t}</a:t></a:r></a:p>`).join("");
    files[`ppt/slides/slide${i + 1}.xml`] =
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="x"><a:txBody>${ps}</a:txBody></p:sld>`;
  });
  return zip(files);
}

function makeDocx(paragraphs) {
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const body = paragraphs.map((p) => `<w:p><w:r><w:t>${p}</w:t></w:r></w:p>`).join("");
  return zip({ "word/document.xml": `<w:document ${ns}><w:body>${body}</w:body></w:document>` });
}

module.exports = { makePdf, makeOdt, makeOds, makeOdp, makePptx, makeDocx };
