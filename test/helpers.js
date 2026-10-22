// What more than one Node suite needs. The two oracles live here: an
// independent reader for what a spreadsheet application would see, and the
// parts themselves for what this writer actually wrote.
import ExcelJS from "exceljs";
import { unzipSync } from "fflate";
import { workbook } from "../lib/index.js";
// Served to the page as well, so it lives in a file with no imports of its own.
import { PNG } from "./fixture.js";

export { PNG };

/**
 * One workbook that reaches every part the writer can emit. Shared so that the
 * suites asserting over a whole document assert over the same one.
 */
export const sample = async () => {
  const wb = workbook({
    title: "Sales",
    creator: "quario",
    subject: "Quarterly figures",
    description: "Every part this writer emits, in one document.",
  });
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  const head = sheet.row([
    { value: "Sales", style: { font: { bold: true, size: 16 } } },
    { value: null },
  ]);
  sheet.merge(head, 1, 2);
  sheet.row([{ value: "Mouse & Pad" }, { value: 1050, style: { numberFormat: "#,##0.00" } }]);
  sheet.row([{ value: new Date(Date.UTC(2024, 1, 29)) }, { value: true }]);
  sheet.row(
    [
      {
        value: "Bordered",
        style: {
          fill: "#eeeeee",
          border: {
            left: { style: "thin" },
            right: { style: "dashed", color: "#369" },
            top: { style: "dotted" },
            bottom: null,
          },
          alignment: { horizontal: "center", vertical: "top", wrapText: true },
        },
      },
      {
        value: [
          { text: "rich " },
          {
            text: "runs",
            font: {
              name: "Georgia",
              italic: true,
              underline: true,
              strikethrough: true,
              color: "#cc0000",
            },
          },
        ],
      },
    ],
    { level: 1, hidden: true },
  );
  const total = sheet.row([{ value: "total" }, { value: 2100 }], { collapsed: true });
  sheet.freeze(1);
  sheet.widths([18, null, 7]);
  sheet.print({
    margin: { top: 36, right: 18, bottom: 36, left: 18 },
    size: "A4",
    orientation: "landscape",
    fit: true,
    titles: 1,
    header: { left: "Report", center: [{ field: "page" }, " of ", { field: "pages" }] },
    footer: [{ text: "draft", bold: true, size: 9 }],
    firstHeader: "Sales",
    firstFooter: { right: "quario" },
  });
  sheet.filter({ top: 1, left: 1, bottom: total, right: 2 });
  sheet.place(logo, { row: 1, width: 120, height: 40 });
  sheet.link(2, 1, { url: "https://example.test/mouse" });
  sheet.link(3, 1, { location: "'Report'!A1" });
  return wb.bytes();
};

/**
 * The package's parts, by name. Names come back in the order they were
 * written, which is what pins the part sequence.
 * @param {Uint8Array} bytes
 */
export const parts = (bytes) => unzipSync(bytes);

/**
 * The bytes read back through an independent implementation. Nothing else in
 * this suite proves the file is one a foreign reader accepts.
 * @param {Uint8Array} bytes
 */
export const book = async (bytes) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  return workbook;
};

/**
 * Every part's text, decoded. The media parts are binary and excluded.
 * @param {Uint8Array} bytes
 */
export const xml = (bytes) => {
  const decoder = new TextDecoder();
  const out = {};
  for (const [name, body] of Object.entries(parts(bytes)))
    if (name.endsWith(".xml") || name.endsWith(".rels")) out[name] = decoder.decode(body);
  return out;
};
