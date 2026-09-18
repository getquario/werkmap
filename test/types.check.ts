// The hand-written declarations, exercised the way a consumer in TypeScript
// would. `tsc` is the assertion; nothing here runs.
import { workbook } from "../lib/index.js";
import type { Cell, Font, Placement, Run, Sheet, Style, Workbook } from "../lib/index.js";

const book: Workbook = workbook({ title: "Sales", creator: "quario" });
const empty: Workbook = workbook();

const font: Font = { name: "Calibri", size: 11, bold: true, strikethrough: true, color: "#008000" };
const style: Style = {
  font,
  fill: "#eeeeee",
  border: { top: { style: "thin", color: "#999999" }, bottom: null },
  alignment: { horizontal: "center", vertical: "top", wrapText: true },
  numberFormat: "#,##0.00",
};

const runs: Run[] = [{ text: "Total " }, { text: "1050", font }];
const cells: (Cell | null)[] = [
  { value: "Product", style },
  { value: 1050 },
  { value: true },
  { value: new Date() },
  { value: runs },
  { value: null, style },
  null,
];

const sheet: Sheet = book.sheet("Report");
const name: string = sheet.name;
const row: number = sheet.row(cells);
sheet.merge(row, 1, 2);
sheet.freeze(row);
sheet.widths([12.5, null, 30]);
sheet.widths([]);

const logo: number = book.image(new Uint8Array([0x89, 0x50]), "png");
const at: Placement = { row, col: 1, width: 120, height: 40 };
sheet.place(logo, at);

const bytes: Promise<Uint8Array> = book.bytes();

export { bytes, empty, name };
