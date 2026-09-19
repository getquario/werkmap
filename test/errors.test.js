// Strictness belongs at the format edge: the writer's one job is never to
// emit a file a reader offers to repair, so everything it cannot represent
// throws at the boundary rather than being quietly dropped. Native
// `TypeError` and `RangeError`, no custom class, no error codes.
import assert from "node:assert/strict";
import test from "node:test";
import { workbook } from "../lib/index.js";
import { PNG } from "./helpers.js";

const sheetOf = () => workbook().sheet("Report");

test("a workbook refuses metadata that is not an object", () => {
  assert.throws(() => workbook(null), TypeError);
  assert.throws(() => workbook("Sales"), TypeError);
});

test("a metadata property must be a string", () => {
  assert.throws(() => workbook({ title: 7 }), { name: "TypeError", message: /title/ });
  assert.throws(() => workbook({ creator: {} }), { name: "TypeError", message: /creator/ });
  assert.throws(() => workbook({ subject: [] }), { name: "TypeError", message: /subject/ });
  assert.throws(() => workbook({ description: false }), {
    name: "TypeError",
    message: /description/,
  });
});

test("a workbook with no sheet has nothing to write", async () => {
  await assert.rejects(() => workbook().bytes(), RangeError);
});

test("a sheet name Excel refuses is refused here", () => {
  const wb = workbook();
  assert.throws(() => wb.sheet(""), RangeError);
  assert.throws(() => wb.sheet(7), RangeError);
  assert.throws(() => wb.sheet("x".repeat(32)), { name: "RangeError", message: /31/ });
  for (const bad of ["a:b", "a\\b", "a/b", "a?b", "a*b", "a[b", "a]b"])
    assert.throws(() => wb.sheet(bad), { name: "RangeError", message: /cannot contain/ });

  wb.sheet("Report");
  assert.throws(() => wb.sheet("Report"), { name: "RangeError", message: /already exists/ });
});

test("a row is an array of cells", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row("Product"), TypeError);
  assert.throws(() => sheet.row([7]), { name: "TypeError", message: /cell 1/ });
  assert.throws(() => sheet.row(Array.from({ length: 16_385 }, () => null)), {
    name: "RangeError",
    message: /16384/,
  });
});

test("a value the format cannot carry throws", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([{ value: Number.NaN }]), RangeError);
  assert.throws(() => sheet.row([{ value: Number.POSITIVE_INFINITY }]), RangeError);
  assert.throws(() => sheet.row([{ value: new Date("nonsense") }]), {
    name: "RangeError",
    message: /valid Date/,
  });
  assert.throws(() => sheet.row([{ value: { product: "Laptop" } }]), {
    name: "TypeError",
    message: /unsupported cell value/,
  });
});

test("a rich-text array must hold runs", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([{ value: [] }]), { name: "RangeError", message: /at least one/ });
  assert.throws(() => sheet.row([{ value: ["Total"] }]), {
    name: "TypeError",
    message: /\{ text, font \} run/,
  });
  assert.throws(() => sheet.row([{ value: [{ text: 7 }] }]), {
    name: "TypeError",
    message: /text/,
  });
});

test("a style must be an object, and its parts must be what they claim", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([{ value: 1, style: "bold" }]), TypeError);
  assert.throws(() => sheet.row([{ value: 1, style: { numberFormat: 7 } }]), {
    name: "TypeError",
    message: /numberFormat/,
  });
  assert.throws(() => sheet.row([{ value: 1, style: { font: { size: "big" } } }]), {
    name: "RangeError",
    message: /finite number/,
  });
  assert.throws(() => sheet.row([{ value: 1, style: { font: { name: 7 } } }]), {
    name: "TypeError",
    message: /font.name/,
  });
});

test("a colour is #rrggbb and nothing else", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([{ value: 1, style: { fill: "red" } }]), {
    name: "TypeError",
    message: /#rrggbb/,
  });
  assert.throws(() => sheet.row([{ value: 1, style: { fill: "#ff00" } }]), TypeError);
  assert.throws(() => sheet.row([{ value: 1, style: { fill: 0xff0000 } }]), TypeError);
  assert.throws(() => sheet.row([{ value: 1, style: { font: { color: "#gggggg" } } }]), TypeError);
  assert.throws(
    () => sheet.row([{ value: 1, style: { border: { top: { style: "thin", color: "x" } } } }]),
    TypeError,
  );
});

test("a border takes one of the three styles the vocabulary names", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([{ value: 1, style: { border: { top: { style: "double" } } } }]), {
    name: "RangeError",
    message: /thin, dashed or dotted/,
  });
  assert.throws(() => sheet.row([{ value: 1, style: { border: { top: {} } } }]), RangeError);
});

test("an alignment takes a name the format knows", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([{ value: 1, style: { alignment: { horizontal: "centre" } } }]), {
    name: "RangeError",
    message: /horizontal/,
  });
  assert.throws(() => sheet.row([{ value: 1, style: { alignment: { vertical: "middle" } } }]), {
    name: "RangeError",
    message: /vertical/,
  });
});

test("a merge names a range that exists and overlaps nothing", () => {
  const sheet = sheetOf();
  const row = sheet.row([null, null, null, null]);

  assert.throws(() => sheet.merge(row, 1, 1), { name: "RangeError", message: /2 or more/ });
  assert.throws(() => sheet.merge(row, 1, 1.5), RangeError);
  assert.throws(() => sheet.merge(row + 1, 1, 2), {
    name: "RangeError",
    message: /does not exist/,
  });
  assert.throws(() => sheet.merge(0, 1, 2), RangeError);
  assert.throws(() => sheet.merge(row, 0, 2), RangeError);
  assert.throws(() => sheet.merge(row, 16_384, 2), { name: "RangeError", message: /past column/ });

  sheet.merge(row, 1, 2);
  assert.throws(() => sheet.merge(row, 2, 2), { name: "RangeError", message: /overlaps/ });
  assert.throws(() => sheet.merge(row, 1, 3), { name: "RangeError", message: /overlaps/ });
});

test("a freeze takes a row count", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.freeze(-1), RangeError);
  assert.throws(() => sheet.freeze(1.5), RangeError);
  assert.throws(() => sheet.freeze("1"), RangeError);
  assert.throws(() => sheet.freeze(1_048_576), RangeError);
});

test("a width list is an array of widths the format can carry", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.widths(12), TypeError);
  assert.throws(() => sheet.widths(null), { name: "TypeError", message: /an array of widths/ });
  assert.throws(() => sheet.widths(Array.from({ length: 16_385 }, () => 10)), {
    name: "RangeError",
    message: /at most 16384 columns/,
  });
  assert.throws(() => sheet.widths([10, "wide"]), {
    name: "RangeError",
    message: /column 2: expected a finite number/,
  });
  assert.throws(() => sheet.widths([Number.POSITIVE_INFINITY]), RangeError);
  assert.throws(() => sheet.widths([10, -1]), {
    name: "RangeError",
    message: /column 2 expected a width above 0 and at most 255, got -1/,
  });
  // Zero is the back door to a hidden column, which is not this surface.
  assert.throws(() => sheet.widths([0]), {
    name: "RangeError",
    message: /null leaves a column unset/,
  });
  assert.throws(() => sheet.widths([255.5]), { name: "RangeError", message: /at most 255/ });
});

test("an image is bytes in a format the writer knows", () => {
  const wb = workbook();
  assert.throws(() => wb.image("logo.png", "png"), { name: "TypeError", message: /Uint8Array/ });
  assert.throws(() => wb.image(PNG, "gif"), { name: "TypeError", message: /png or jpeg/ });
});

test("a placement names an image and a cell that exist", () => {
  const wb = workbook();
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "anchor" }]);

  assert.throws(() => sheet.place(1, { row: 1, width: 10, height: 10 }), {
    name: "TypeError",
    message: /no image with id 1/,
  });
  assert.throws(() => sheet.place("logo", { row: 1, width: 10, height: 10 }), TypeError);
  assert.throws(() => sheet.place(logo, null), TypeError);
  assert.throws(() => sheet.place(logo, { row: 0, width: 10, height: 10 }), RangeError);
  assert.throws(() => sheet.place(logo, { row: 1, col: 0, width: 10, height: 10 }), RangeError);
  assert.throws(() => sheet.place(logo, { row: 1, width: "10", height: 10 }), RangeError);
  assert.throws(() => sheet.place(logo, { row: 1, width: 10, height: Number.NaN }), RangeError);
  assert.throws(() => sheet.place(logo, { row: 1, width: 0, height: 10 }), {
    name: "RangeError",
    message: /positive size/,
  });
  assert.throws(() => sheet.place(logo, { row: 1, width: 10, height: -1 }), RangeError);
});

test("a row's outline level is an integer Excel can draw", () => {
  const sheet = sheetOf();
  assert.throws(() => sheet.row([], "1"), { name: "TypeError", message: /row: options/ });
  for (const level of [-1, 8, 1.5, "1"])
    assert.throws(() => sheet.row([], { level }), {
      name: "RangeError",
      message: /row: level: expected an integer between 0 and 7/,
    });
});
