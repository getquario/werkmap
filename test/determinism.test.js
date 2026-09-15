// The headline promise: the same calls produce the same bytes on the same
// runtime. It is tested by self-comparison rather than against a stored
// golden, because a golden would pin whichever zlib the runner happens to
// ship — "Node 22" is not one zlib, and a patch release can carry a new fork
// revision — while the promise itself holds either way.
import assert from "node:assert/strict";
import test from "node:test";
import { workbook } from "../lib/index.js";
import { PNG, book, parts } from "./helpers.js";

const build = async () => {
  const wb = workbook({ title: "Sales", creator: "quario" });
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  const head = sheet.row([
    { value: "Sales", style: { font: { bold: true, size: 16 } } },
    { value: null },
  ]);
  sheet.merge(head, 1, 2);
  sheet.row([{ value: "Mouse & Pad" }, { value: 1050, style: { numberFormat: "#,##0.00" } }]);
  sheet.row([{ value: new Date(Date.UTC(2024, 1, 29)) }, { value: true }]);
  sheet.freeze(1);
  sheet.place(logo, { row: 1, width: 120, height: 40 });
  return wb.bytes();
};

test("the same calls produce the same bytes", async () => {
  const first = await build();
  const second = await build();
  assert.deepEqual(first, second);
});

test("one workbook can be written more than once", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: "once" }]);
  assert.deepEqual(await wb.bytes(), await wb.bytes(), "bytes() does not seal the document");
});

test("the archive unzips, and every part is well-formed XML a reader accepts", async () => {
  const bytes = await build();
  const entries = parts(bytes);
  const decoder = new TextDecoder();

  for (const [name, body] of Object.entries(entries)) {
    if (name.endsWith(".png")) continue;
    const text = decoder.decode(body);
    assert.ok(text.startsWith('<?xml version="1.0"'), `${name} opens with a declaration`);
    assert.ok(!text.includes("undefined"), `${name} carries no stringified undefined`);
  }

  // exceljs parses every part it loads, so a malformed one fails here rather
  // than in a spreadsheet application.
  const read = await book(bytes);
  assert.equal(read.getWorksheet("Report").getCell("A1").value, "Sales");
});

test("the parts are written in a fixed order", async () => {
  assert.deepEqual(Object.keys(parts(await build())), [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/core.xml",
    "docProps/app.xml",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
    "xl/media/image1.png",
  ]);
});

/**
 * The local file headers, read straight out of the container: name, the
 * compression method, and the DOS timestamp.
 * @param {Uint8Array} bytes
 */
const headers = (bytes) => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out = [];
  let at = 0;
  while (view.getUint32(at, true) === 0x04034b50) {
    const method = view.getUint16(at + 8, true);
    const time = view.getUint16(at + 10, true);
    const date = view.getUint16(at + 12, true);
    const packed = view.getUint32(at + 18, true);
    const length = view.getUint16(at + 26, true);
    const name = new TextDecoder().decode(bytes.subarray(at + 30, at + 30 + length));
    out.push({ name, method, time, date });
    at += 30 + length + packed;
  }
  return out;
};

test("media is stored and XML is deflated", async () => {
  const entries = headers(await build());
  assert.equal(entries.length, 13);
  for (const entry of entries)
    assert.equal(
      entry.method,
      entry.name.endsWith(".png") ? 0 : 8,
      `${entry.name} uses the wrong compression method`,
    );
});

test("every entry carries the pinned ZIP epoch", async () => {
  for (const entry of headers(await build())) {
    assert.equal(entry.time, 0, `${entry.name} carries a time`);
    assert.equal(entry.date, (1 << 5) | 1, `${entry.name} carries a date other than 1980-01-01`);
  }
});

test("stored media is byte-identical to what the caller handed over", async () => {
  const wb = workbook();
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "anchor" }]);
  sheet.place(logo, { row: 1, width: 10, height: 10 });
  assert.deepEqual(parts(await wb.bytes())["xl/media/image1.png"], PNG);
});
