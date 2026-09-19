// What a foreign reader makes of what this writer wrote. Every assertion here
// goes through exceljs rather than through the XML, so it reads what a
// spreadsheet application would read.
import assert from "node:assert/strict";
import test from "node:test";
import { workbook } from "../lib/index.js";
import { PNG, book, parts, xml } from "./helpers.js";

test("a workbook carries its metadata and a pinned clock", async () => {
  const wb = workbook({
    title: "Sales",
    creator: "quario",
    subject: "Q3",
    description: "Rendered with an unlicensed copy",
  });
  wb.sheet("Report").row([{ value: "Hello" }]);
  const read = await book(await wb.bytes());

  assert.equal(read.title, "Sales");
  assert.equal(read.creator, "quario");
  assert.equal(read.subject, "Q3");
  assert.equal(read.description, "Rendered with an unlicensed copy");
  assert.equal(read.created.getTime(), 0, "created is pinned to the epoch");
  assert.equal(read.modified.getTime(), 0, "and so is modified");
});

test("metadata is optional, and a workbook needs no arguments at all", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: 1 }]);
  const read = await book(await wb.bytes());
  assert.equal(read.title, undefined);
  assert.equal(read.getWorksheet("Report").getCell("A1").value, 1);
});

test("cells keep their types", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "text" }, { value: 1234.5 }, { value: true }, { value: false }]);
  sheet.row([{ value: new Date(Date.UTC(2024, 1, 29, 12)) }, { value: null }, null]);
  const ws = (await book(await wb.bytes())).getWorksheet("Report");

  assert.equal(ws.getCell("A1").value, "text");
  assert.equal(ws.getCell("B1").value, 1234.5);
  assert.equal(ws.getCell("C1").value, true);
  assert.equal(ws.getCell("D1").value, false);
  assert.deepEqual(ws.getCell("A2").value, new Date(Date.UTC(2024, 1, 29, 12)));
  assert.equal(ws.getCell("B2").value, null);
});

test("a date before the phantom leap day converts correctly", async () => {
  const wb = workbook();
  // 1900-02-28 is serial 59; Excel's phantom 1900-02-29 sits at 60, which the
  // 1899-12-30 epoch absorbs.
  wb.sheet("Report").row([{ value: new Date(Date.UTC(1900, 1, 28)) }]);
  const sheet = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  assert.match(sheet, /<v>59<\/v>/);
});

test("a date gets a date format unless the caller names one", async () => {
  const wb = workbook();
  const when = new Date(Date.UTC(2024, 1, 29));
  wb.sheet("Report").row([
    { value: when },
    { value: when, style: null },
    { value: when, style: { font: { bold: true } } },
    { value: when, style: { numberFormat: null } },
    { value: when, style: { numberFormat: "dd mmm yyyy" } },
  ]);
  const ws = (await book(await wb.bytes())).getWorksheet("Report");

  for (const at of ["A1", "B1", "C1", "D1"])
    assert.equal(ws.getCell(at).numFmt, "mm-dd-yy", `${at} took the short-date built-in`);
  assert.equal(ws.getCell("E1").numFmt, "dd mmm yyyy", "a named format wins");
  assert.deepEqual(ws.getCell("A1").value, when, "and the value is still a date");
});

test("a style that is not an object is refused even on a date cell", () => {
  const sheet = workbook().sheet("Report");
  assert.throws(() => sheet.row([{ value: new Date(0), style: "bold" }]), TypeError);
});

test("a merge can reach past the widest row", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  const row = sheet.row([{ value: "wide" }]);
  sheet.merge(row, 1, 3);
  assert.match(
    xml(await wb.bytes())["xl/worksheets/sheet1.xml"],
    /<dimension ref="A1:C1"\/>/,
    "the merged range widens the sheet",
  );
});

test("a very large number writes an uppercase exponent", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: 1e21 }]);
  const sheet = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  assert.match(sheet, /<v>1E\+21<\/v>/, "Excel writes E, not e");
});

test("rich text is a bare array, and every run carries its font in full", async () => {
  const wb = workbook();
  wb.sheet("Report").row([
    {
      value: [
        { text: "Total " },
        { text: "bold", font: { bold: true, color: "#b91c1c" } },
        { text: " and plain", font: null },
      ],
    },
  ]);
  const ws = (await book(await wb.bytes())).getWorksheet("Report");
  const value = ws.getCell("A1").value;

  assert.equal(value.richText.length, 3);
  assert.equal(value.richText[0].text, "Total ");
  assert.equal(value.richText[1].text, "bold");
  assert.equal(value.richText[1].font.bold, true);
  assert.equal(value.richText[1].font.color.argb, "FFB91C1C");
  assert.equal(value.richText[2].text, " and plain");
});

test("the whole style vocabulary reaches the cell", async () => {
  const wb = workbook();
  wb.sheet("Report").row([
    {
      value: "styled",
      style: {
        font: {
          name: "Times New Roman",
          size: 16,
          bold: true,
          italic: true,
          underline: true,
          strikethrough: true,
          color: "#008000",
        },
        fill: "#eee",
        border: {
          top: { style: "thin", color: "#999999" },
          bottom: { style: "dashed" },
          left: null,
        },
        alignment: { horizontal: "center", vertical: "top", wrapText: true },
        numberFormat: "#,##0.00",
      },
    },
  ]);
  const cell = (await book(await wb.bytes())).getWorksheet("Report").getCell("A1");

  assert.equal(cell.font.name, "Times New Roman");
  assert.equal(cell.font.size, 16);
  assert.equal(cell.font.bold, true);
  assert.equal(cell.font.italic, true);
  assert.equal(cell.font.underline, true);
  assert.equal(cell.font.strike, true);
  assert.equal(cell.font.color.argb, "FF008000");
  assert.equal(cell.fill.fgColor.argb, "FFEEEEEE", "the #rgb shorthand expands");
  assert.equal(cell.border.top.style, "thin");
  assert.equal(cell.border.top.color.argb, "FF999999");
  assert.equal(cell.border.bottom.style, "dashed");
  assert.equal(cell.alignment.horizontal, "center");
  assert.equal(cell.alignment.vertical, "top");
  assert.equal(cell.alignment.wrapText, true);
  assert.equal(cell.numFmt, "#,##0.00");
});

test("a dotted border and the remaining alignments are accepted", async () => {
  const wb = workbook();
  wb.sheet("Report").row([
    {
      value: null,
      style: {
        border: { right: { style: "dotted" } },
        alignment: { horizontal: "justify", vertical: "justify" },
      },
    },
  ]);
  const cell = (await book(await wb.bytes())).getWorksheet("Report").getCell("A1");
  assert.equal(cell.border.right.style, "dotted");
  assert.equal(cell.alignment.horizontal, "justify");
});

test("a built-in format code is written as its id, a custom one is interned", async () => {
  const wb = workbook();
  wb.sheet("Report").row([
    { value: 1, style: { numberFormat: "0.00%" } },
    { value: 2, style: { numberFormat: "dd mmm yyyy" } },
    { value: 3, style: { numberFormat: "dd mmm yyyy" } },
    { value: 4, style: { numberFormat: "0.000" } },
  ]);
  const bytes = await wb.bytes();
  const styles = xml(bytes)["xl/styles.xml"];

  assert.match(styles, /numFmtId="164" formatCode="dd mmm yyyy"/);
  assert.match(styles, /numFmtId="165" formatCode="0.000"/);
  assert.doesNotMatch(styles, /formatCode="0.00%"/, "a built-in needs no entry of its own");
  assert.match(styles, /<numFmts count="2">/, "the repeated code interned once");

  const ws = (await book(bytes)).getWorksheet("Report");
  assert.equal(ws.getCell("A1").numFmt, "0.00%");
  assert.equal(ws.getCell("B1").numFmt, "dd mmm yyyy");
});

test("identical styles collapse to one entry", async () => {
  const wb = workbook();
  const style = { font: { bold: true }, fill: "#ff0000" };
  const sheet = wb.sheet("Report");
  sheet.row([
    { value: "a", style },
    { value: "b", style: { ...style } },
  ]);
  sheet.row([{ value: "c", style: { font: { bold: true }, fill: "#ff0000" } }]);
  const styles = xml(await wb.bytes())["xl/styles.xml"];

  assert.match(styles, /<fonts count="2">/, "one default font and one bold one");
  assert.match(styles, /<fills count="3">/, "the two reserved fills and one red");
  assert.match(styles, /<cellXfs count="2">/, "one default format and one styled");
});

test("repeated text interns once", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "Product" }, { value: "Product" }]);
  sheet.row([{ value: "Product" }]);
  const shared = xml(await wb.bytes())["xl/sharedStrings.xml"];
  assert.match(shared, /count="1" uniqueCount="1"/);
});

test("markup and control characters reach the cell as inert text", async () => {
  const wb = workbook();
  wb.sheet("Report").row([
    { value: '<script>alert("x")</script> & co' },
    { value: "bell\u0007 and tab\there" },
    { value: "a literal _x0041_ sequence" },
    { value: "lone \ud800 surrogate" },
  ]);
  const ws = (await book(await wb.bytes())).getWorksheet("Report");

  assert.equal(ws.getCell("A1").value, '<script>alert("x")</script> & co');
  assert.equal(
    ws.getCell("B1").value,
    "bell and tab\there",
    "the C0 control is stripped, TAB is not",
  );
  assert.equal(ws.getCell("C1").value, "a literal _x0041_ sequence", "the escape survives as text");
  assert.equal(ws.getCell("D1").value, "lone  surrogate", "the unpaired surrogate is stripped");
});

test("leading and trailing whitespace survives", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: "  padded  " }]);
  const ws = (await book(await wb.bytes())).getWorksheet("Report");
  assert.equal(ws.getCell("A1").value, "  padded  ");
});

test("a merged range spans the columns it names", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  const row = sheet.row([{ value: "Title" }, null, null]);
  sheet.merge(row, 1, 3);
  sheet.row([{ value: "below" }]);
  const ws = (await book(await wb.bytes())).getWorksheet("Report");

  assert.ok(ws.getCell("A1").isMerged);
  assert.ok(ws.getCell("C1").isMerged);
  assert.ok(!ws.getCell("A2").isMerged);
});

test("two merges on one row sit side by side", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  const row = sheet.row([null, null, null, null]);
  sheet.merge(row, 1, 2);
  sheet.merge(row, 3, 2);
  const sheetXml = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  assert.match(sheetXml, /<mergeCells count="2">/);
  assert.match(sheetXml, /ref="C1:D1"/);
});

test("a frozen pane holds, and zero clears it", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "head" }]);
  sheet.freeze(1);
  let ws = (await book(await wb.bytes())).getWorksheet("Report");
  assert.equal(ws.views[0].state, "frozen");
  assert.equal(ws.views[0].ySplit, 1);

  sheet.freeze(0);
  ws = (await book(await wb.bytes())).getWorksheet("Report");
  assert.notEqual(ws.views[0].state, "frozen");
});

test("an image floats at the cell it anchors to", async () => {
  const wb = workbook();
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "row one" }]);
  sheet.row([{ value: "row two" }]);
  sheet.place(logo, { row: 2, col: 2, width: 120, height: 40 });

  const bytes = await wb.bytes();
  const read = await book(bytes);
  const [drawing] = read.getWorksheet("Report").getImages();

  assert.equal(drawing.range.tl.nativeRow, 1, "1-based at the surface, 0-based in the XML");
  assert.equal(drawing.range.tl.nativeCol, 1);
  assert.equal(read.model.media[drawing.imageId].extension, "png");
  assert.match(
    xml(bytes)["xl/drawings/drawing1.xml"],
    /cx="1143000" cy="381000"/,
    "CSS pixels at 96 dpi become EMU",
  );
});

test("`col` defaults to the first column", async () => {
  const wb = workbook();
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "anchor" }]);
  sheet.place(logo, { row: 1, width: 10, height: 10 });
  const [drawing] = (await book(await wb.bytes())).getWorksheet("Report").getImages();
  assert.equal(drawing.range.tl.nativeCol, 0);
});

test("identical image bytes deduplicate, and a format change does not", async () => {
  const wb = workbook();
  const first = wb.image(PNG, "png");
  const twice = wb.image(PNG, "png");
  const again = wb.image(new Uint8Array(PNG), "png");
  const asJpeg = wb.image(PNG, "jpeg");
  const shorter = wb.image(PNG.slice(0, 20), "png");
  const different = wb.image(new Uint8Array(PNG.length).fill(7), "png");

  assert.equal(twice, first, "the same array returns the id already issued");
  assert.equal(again, first, "and so do equal bytes in another array");
  assert.notEqual(asJpeg, first, "a different format is a different entry");
  assert.notEqual(shorter, first);
  assert.notEqual(different, first);

  const sheet = wb.sheet("Report");
  sheet.row([{ value: "anchor" }]);
  sheet.place(first, { row: 1, width: 10, height: 10 });
  const names = Object.keys(parts(await wb.bytes()));
  assert.ok(names.includes("xl/media/image1.png"));
  assert.ok(
    names.includes("xl/media/image2.jpeg"),
    "the same bytes under another format are their own entry",
  );
});

test("one image placed twice on a sheet is one relationship", async () => {
  const wb = workbook();
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "a" }]);
  sheet.row([{ value: "b" }]);
  sheet.place(logo, { row: 1, width: 10, height: 10 });
  sheet.place(logo, { row: 2, width: 10, height: 10 });

  const rels = xml(await wb.bytes())["xl/drawings/_rels/drawing1.xml.rels"];
  assert.equal(rels.match(/<Relationship /g).length, 1);
});

test("several sheets keep their own rows, drawings and order", async () => {
  const wb = workbook();
  const logo = wb.image(PNG, "png");
  const first = wb.sheet("First");
  const second = wb.sheet("Second");
  const third = wb.sheet("Third");

  first.row([{ value: "one" }]);
  second.row([{ value: "two" }]);
  second.place(logo, { row: 1, width: 10, height: 10 });
  third.row([{ value: "three" }]);
  third.place(logo, { row: 1, width: 10, height: 10 });

  const bytes = await wb.bytes();
  const read = await book(bytes);
  assert.deepEqual(
    read.worksheets.map((each) => each.name),
    ["First", "Second", "Third"],
  );
  assert.equal(read.getWorksheet("Second").getCell("A1").value, "two");
  assert.equal(read.getWorksheet("First").getImages().length, 0);
  assert.equal(read.getWorksheet("Third").getImages().length, 1);

  const names = Object.keys(parts(bytes));
  assert.ok(names.includes("xl/drawings/drawing2.xml"), "drawings number from one, not by sheet");
  assert.ok(
    !names.includes("xl/worksheets/_rels/sheet1.xml.rels"),
    "a sheet with no drawing has no rels part",
  );
});

test("a sheet with no rows is still a sheet", async () => {
  const wb = workbook();
  wb.sheet("Empty");
  const bytes = await wb.bytes();
  assert.match(xml(bytes)["xl/worksheets/sheet1.xml"], /<dimension ref="A1:A1"\/>/);
  assert.equal((await book(bytes)).getWorksheet("Empty").name, "Empty");
});

test("column widths reach a reader in the format's own unit", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "Product" }, { value: 1 }, { value: 2 }]);
  sheet.widths([12.5, null, 255]);
  const bytes = await wb.bytes();

  const ws = (await book(bytes)).getWorksheet("Report");
  assert.equal(ws.getColumn(1).width, 12.5);
  assert.equal(ws.getColumn(3).width, 255);
  // An unset column carries nothing, so the reader keeps its own default.
  assert.equal(ws.getColumn(2).width, undefined);
});

test("a width is written verbatim, one `col` apiece, before the rows", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  // Nothing rounds it: the unit needed no conversion, so the writer owns no
  // arithmetic here and has nothing to round.
  sheet.widths([0.1 + 0.2, 8]);
  const written = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];

  assert.match(
    written,
    /<\/sheetViews><cols><col min="1" max="1" width="0.30000000000000004" customWidth="1"\/><col min="2" max="2" width="8" customWidth="1"\/><\/cols><sheetData>/,
    "`cols` sits between `sheetViews` and `sheetData`, one element per set column",
  );
});

test("a sheet nobody sized carries no cols element", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: 1 }]);
  assert.doesNotMatch(xml(await wb.bytes())["xl/worksheets/sheet1.xml"], /<cols>/);
});

test("a width list replaces, and an all-empty list writes nothing", async () => {
  const wb = workbook();
  const one = wb.sheet("Kept");
  one.row([{ value: 1 }]);
  one.widths([40, 40]);
  // The second call is the sheet's widths now -- column 2 goes back to unset.
  one.widths([12]);

  const two = wb.sheet("Cleared");
  two.row([{ value: 1 }]);
  two.widths([40]);
  two.widths([]);

  const three = wb.sheet("Holes");
  three.row([{ value: 1 }]);
  three.widths([null, null]);

  const bytes = await wb.bytes();
  const read = await book(bytes);
  assert.equal(read.getWorksheet("Kept").getColumn(1).width, 12);
  assert.equal(read.getWorksheet("Kept").getColumn(2).width, undefined);
  assert.doesNotMatch(xml(bytes)["xl/worksheets/sheet2.xml"], /<cols>/);
  assert.doesNotMatch(xml(bytes)["xl/worksheets/sheet3.xml"], /<cols>/);
});

test("sizing a column writes no cell, so the sheet stays as wide as its rows", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  // Sized before any row, and past every row there will be.
  sheet.widths([10, 10, 10, 10]);
  sheet.row([{ value: 1 }]);
  const written = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];

  assert.match(written, /<dimension ref="A1:A1"\/>/);
  assert.match(written, /<col min="4" max="4"/, "the width itself is still written");
});

test("a worksheet carries no print setup unless it asks for one", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: 1 }]);
  const sheet = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  // A reader's own defaults beat this writer's opinion, so nothing is written.
  assert.doesNotMatch(sheet, /<pageMargins|<pageSetup|<sheetPr/);
});

test("print setup writes the parts it was asked for, in the schema's order", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ margin: 54, size: "A4", orientation: "landscape", fit: true });
  const written = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];

  // `sheetPr` opens the worksheet; the two print parts sit after `mergeCells`
  // and before `drawing`.
  assert.match(
    written,
    /<worksheet[^>]*><sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr><dimension/,
  );
  assert.match(
    written,
    /<pageMargins left="0.750" right="0.750" top="0.750" bottom="0.750" header="0.3" footer="0.3"\/>/,
    "points convert to inches, and the two margins this surface omits keep Excel's gap",
  );
  assert.match(
    written,
    /<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"\/>/,
  );
});

test("print setup takes each key on its own, and calls merge", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ size: "legal" });
  sheet.print({ orientation: "portrait" });
  const written = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];

  assert.doesNotMatch(written, /<pageMargins/, "no margin was named");
  assert.doesNotMatch(written, /<sheetPr/, "no fit was named");
  assert.match(written, /<pageSetup paperSize="5" orientation="portrait"\/>/);
});

test("a margin alone writes margins and no page setup", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ margin: 0 });
  const written = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  assert.match(written, /<pageMargins left="0.000"/);
  assert.doesNotMatch(written, /<pageSetup/);
});

test("print titles repeat the top rows, scoped to their own worksheet", async () => {
  const wb = workbook();
  const one = wb.sheet("North");
  one.row([{ value: "Region" }]);
  one.row([{ value: 1 }]);
  one.print({ titles: 1 });
  const two = wb.sheet("South");
  two.row([{ value: 2 }]);
  two.print({ titles: 2 });
  wb.sheet("Quiet").row([{ value: 3 }]);
  const bytes = await wb.bytes();

  // The names sit after `</sheets>`, which is where the schema puts them, and
  // `localSheetId` is the 0-based position in that same list. A sheet that
  // asked for none contributes no name, so the third has none.
  assert.match(
    xml(bytes)["xl/workbook.xml"],
    /<\/sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">'North'!\$1:\$1<\/definedName><definedName name="_xlnm.Print_Titles" localSheetId="1">'South'!\$1:\$2<\/definedName><\/definedNames><\/workbook>/,
  );

  // And an independent reader agrees about which rows repeat. It hands the
  // range back without the absolute markers it parsed, so this is exceljs's
  // spelling of the reference above rather than a second claim about the file.
  const read = await book(bytes);
  assert.equal(read.getWorksheet("North").pageSetup.printTitlesRow, "1:1");
  assert.equal(read.getWorksheet("South").pageSetup.printTitlesRow, "1:2");
  assert.equal(read.getWorksheet("Quiet").pageSetup.printTitlesRow, undefined);

  // A title count is not print setup in the sheet part, so asking for one
  // alone still leaves a reader its own margins and paper.
  assert.doesNotMatch(xml(bytes)["xl/worksheets/sheet1.xml"], /<pageSetup|<pageMargins|<sheetPr/);
});

test("a workbook nobody asked for titles on carries no defined names", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: 1 }]);
  assert.doesNotMatch(xml(await wb.bytes())["xl/workbook.xml"], /definedName/);
});

test("a sheet name carrying an apostrophe survives the reference", async () => {
  // A defined name quotes the sheet, so an apostrophe in the name has to
  // double or the reference ends early. It is XML-escaped after that, like
  // every other name this writer emits.
  const wb = workbook();
  const sheet = wb.sheet("Bob's & Co");
  sheet.row([{ value: 1 }]);
  sheet.print({ titles: 1 });
  const bytes = await wb.bytes();
  assert.match(xml(bytes)["xl/workbook.xml"], />'Bob''s &amp; Co'!\$1:\$1</);
  assert.equal((await book(bytes)).getWorksheet("Bob's & Co").pageSetup.printTitlesRow, "1:1");
});

test("print titles clear at zero, the way a freeze does", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ titles: 2 });
  sheet.print({ titles: 0 });
  assert.doesNotMatch(xml(await wb.bytes())["xl/workbook.xml"], /definedName/);
});

test("an autofilter covers the range it was given, in the schema's order", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  for (let n = 0; n < 4; n++) sheet.row([{ value: n }, { value: n }, { value: n }]);
  sheet.merge(1, 1, 2);
  sheet.filter({ top: 1, left: 1, bottom: 4, right: 3 });
  const bytes = await wb.bytes();

  // `autoFilter` follows `sheetData` and precedes `mergeCells` in the
  // worksheet's child sequence.
  assert.match(
    xml(bytes)["xl/worksheets/sheet1.xml"],
    /<\/sheetData><autoFilter ref="A1:C4"\/><mergeCells/,
  );
  // And an independent reader finds the same range.
  assert.equal((await book(bytes)).getWorksheet("Report").autoFilter, "A1:C4");
});

test("a worksheet carries no autofilter unless it asks for one", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: 1 }]);
  assert.doesNotMatch(xml(await wb.bytes())["xl/worksheets/sheet1.xml"], /<autoFilter/);
});

test("an autofilter replaces rather than merges, and null clears", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  for (let n = 0; n < 3; n++) sheet.row([{ value: n }, { value: n }]);
  sheet.filter({ top: 1, left: 1, bottom: 3, right: 2 });
  sheet.filter({ top: 2, left: 1, bottom: 3, right: 1 });
  assert.match(xml(await wb.bytes())["xl/worksheets/sheet1.xml"], /<autoFilter ref="A2:A3"\/>/);

  sheet.filter(null);
  assert.doesNotMatch(xml(await wb.bytes())["xl/worksheets/sheet1.xml"], /<autoFilter/);
});

test("an autofilter refuses a range a reader could not carry", () => {
  const sheet = workbook().sheet("Report");
  sheet.row([{ value: 1 }, { value: 2 }]);
  assert.throws(() => sheet.filter("A1:B2"), TypeError);
  assert.throws(() => sheet.filter({ top: 0, left: 1, bottom: 1, right: 1 }), {
    name: "RangeError",
    message: /filter: top/,
  });
  assert.throws(() => sheet.filter({ top: 2, left: 1, bottom: 1, right: 1 }), {
    name: "RangeError",
    message: /filter: bottom 1 is above top 2/,
  });
  assert.throws(() => sheet.filter({ top: 1, left: 2, bottom: 1, right: 1 }), {
    name: "RangeError",
    message: /filter: right 1 is left of left 2/,
  });
  // A filter over cells nobody wrote is a caller's bug, the way a merge over
  // one is -- and the column rule is the row rule, so both ends are checked.
  assert.throws(() => sheet.filter({ top: 1, left: 1, bottom: 9, right: 1 }), {
    name: "RangeError",
    message: /filter: row 9 does not exist yet \(the sheet has 1\)/,
  });
  assert.throws(() => sheet.filter({ top: 1, left: 1, bottom: 1, right: 9 }), {
    name: "RangeError",
    message: /filter: column 9 does not exist yet \(the sheet has 2\)/,
  });
});

test("print setup refuses what a reader could not carry", () => {
  const sheet = workbook().sheet("Report");
  assert.throws(() => sheet.print(null), TypeError);
  assert.throws(() => sheet.print("A4"), TypeError);
  assert.throws(() => sheet.print({ margin: "wide" }), RangeError);
  assert.throws(() => sheet.print({ margin: -1 }), { name: "RangeError", message: /negative/ });
  assert.throws(() => sheet.print({ size: "A2" }), {
    name: "RangeError",
    message: /unknown paper size/,
  });
  assert.throws(() => sheet.print({ orientation: "sideways" }), {
    name: "RangeError",
    message: /portrait or landscape/,
  });
  assert.throws(() => sheet.print({ titles: 1.5 }), {
    name: "RangeError",
    message: /print: titles/,
  });
  // 0 is the clear, so the floor a caller is told about is 0 and not 1.
  assert.throws(() => sheet.print({ titles: -1 }), {
    name: "RangeError",
    message: /titles expected 0, or a row count between 1 and 1048576, got -1/,
  });
  assert.throws(() => sheet.print({ titles: 1_048_577 }), {
    name: "RangeError",
    message: /print: titles/,
  });
});

test("columns past Z carry their letters", async () => {
  const wb = workbook();
  const cells = Array.from({ length: 703 }, (_, index) => ({ value: index + 1 }));
  wb.sheet("Report").row(cells);
  const sheet = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];

  assert.match(sheet, /r="Z1"/);
  assert.match(sheet, /r="AA1"/);
  assert.match(sheet, /r="ZZ1"/);
  assert.match(sheet, /r="AAA1"/);
  assert.match(sheet, /<dimension ref="A1:AAA1"\/>/);
});

test("an empty cell is written only when it carries a style", async () => {
  const wb = workbook();
  wb.sheet("Report").row([{ value: null }, { value: null, style: { fill: "#ff0000" } }]);
  const sheet = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];

  assert.doesNotMatch(sheet, /r="A1"/, "an unstyled empty cell is nothing at all");
  assert.match(sheet, /<c r="B1" s="1"\/>/);
});

test("a sheet name reaches the workbook escaped", async () => {
  const wb = workbook();
  wb.sheet("A & B").row([{ value: 1 }]);
  assert.match(xml(await wb.bytes())["xl/workbook.xml"], /name="A &amp; B"/);
});

test("outline levels reach a reader as row groups, summarised below", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "North" }], { level: 1 });
  sheet.row([{ value: "Laptop" }, { value: 1000 }], { level: 2 });
  sheet.row([{ value: "Mouse" }, { value: 25 }], { level: 2 });
  sheet.row([{ value: "Subtotal" }, { value: 1025 }], { level: 1 });
  sheet.row([{ value: "Total" }, { value: 1025 }]);
  const bytes = await wb.bytes();

  const ws = (await book(bytes)).getWorksheet("Report");
  assert.deepEqual(
    [1, 2, 3, 4, 5].map((n) => ws.getRow(n).outlineLevel),
    [1, 2, 2, 1, 0],
    "each row carries the level it was given, and none is 0 by mistake",
  );
  assert.equal(ws.properties.outlineLevelRow, 2, "the sheet states its deepest level");
  // `sheetFormatPr` sits between `sheetViews` and `sheetData` in the
  // worksheet's child sequence, and a level of 0 writes no attribute at all.
  const part = xml(bytes)["xl/worksheets/sheet1.xml"];
  assert.match(
    part,
    /<\/sheetViews><sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"\/><sheetData>/,
  );
  assert.match(part, /<row r="5"><c /);
});

test("a sheet with no outline states no row height and no level count", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.row([{ value: 2 }], {});
  sheet.row([{ value: 3 }], { level: 0 });
  sheet.row([{ value: 4 }], { level: null });
  const part = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  assert.doesNotMatch(part, /sheetFormatPr|outlineLevel/);
});

test("a print header and footer reach a reader as page furniture, left-aligned", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ header: "Sales & Marketing\nQ3", footer: "Confidential <internal>" });
  const bytes = await wb.bytes();

  const ws = (await book(bytes)).getWorksheet("Report");
  assert.equal(ws.headerFooter.oddHeader, "&LSales && Marketing\nQ3");
  assert.equal(ws.headerFooter.oddFooter, "&LConfidential <internal>");
  // `headerFooter` follows `pageSetup` and precedes `drawing` in the
  // worksheet's child sequence, and the text is escaped for XML as well.
  const part = xml(bytes)["xl/worksheets/sheet1.xml"];
  assert.match(
    part,
    /<headerFooter><oddHeader>&amp;LSales &amp;&amp; Marketing\nQ3<\/oddHeader><oddFooter>&amp;LConfidential &lt;internal&gt;<\/oddFooter><\/headerFooter><\/worksheet>/,
  );
});

test("a header's parts name the page number and count, which the reader fills in", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ footer: ["Page ", { field: "page" }, " of ", { field: "pages" }, " & done"] });
  const ws = (await book(await wb.bytes())).getWorksheet("Report");
  assert.equal(ws.headerFooter.oddFooter, "&LPage &P of &N && done");
});

test("a header alone writes no footer, and merges with the rest of the setup", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({ size: "A4" });
  sheet.print({ header: "Top" });
  const part = xml(await wb.bytes())["xl/worksheets/sheet1.xml"];
  assert.match(
    part,
    /<pageSetup paperSize="9"\/><headerFooter><oddHeader>&amp;LTop<\/oddHeader><\/headerFooter>/,
  );
  assert.doesNotMatch(part, /oddFooter/);
});

test("a header's three sections print where they are named, and the first page may differ", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({
    header: { left: "Acme", right: ["Page ", { field: "page" }] },
    footer: { center: "Confidential" },
    firstHeader: { center: [{ text: "Sales report", bold: true, size: 14 }] },
  });
  const bytes = await wb.bytes();
  const hf = (await book(bytes)).getWorksheet("Report").headerFooter;
  assert.equal(hf.oddHeader, "&LAcme&RPage &P");
  assert.equal(hf.oddFooter, "&CConfidential");
  assert.equal(hf.firstHeader, "&C&B&14Sales report");
  assert.equal(hf.differentFirst, true);
  // The parts sit in the schema's order, and the attribute rides on the
  // element only because a first-page part was named.
  assert.match(
    xml(bytes)["xl/worksheets/sheet1.xml"],
    /<headerFooter differentFirst="1"><oddHeader>[^<]*<\/oddHeader><oddFooter>[^<]*<\/oddFooter><firstHeader>[^<]*<\/firstHeader><\/headerFooter>/,
  );
});

test("a look holds until a part changes it, and starts plain in every section", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: 1 }]);
  sheet.print({
    header: [
      { text: "Bold ", bold: true },
      { text: "still bold, bigger ", bold: true, size: 12 },
      { text: "same size, plain ", size: 12 },
      "and a string is plain too ",
      { text: "small", size: 8 },
    ],
    footer: { left: [{ text: "left", bold: true }], right: [{ text: "right" }] },
  });
  const hf = (await book(await wb.bytes())).getWorksheet("Report").headerFooter;
  assert.equal(
    hf.oddHeader,
    "&L&BBold &12still bold, bigger &Bsame size, plain and a string is plain too &08small",
  );
  assert.equal(
    hf.oddFooter,
    "&L&Bleft&Rright",
    "the right section does not inherit the left's bold",
  );
});
