// What more than one Node suite needs. The two oracles live here: an
// independent reader for what a spreadsheet application would see, and the
// parts themselves for what this writer actually wrote.
import ExcelJS from "exceljs";
import { unzipSync } from "fflate";

// Served to the page as well, so it lives in a file with no imports of its own.
export { PNG } from "./fixture.js";

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
