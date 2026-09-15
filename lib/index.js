// werkmap — a write-only OOXML (.xlsx) writer.
//
// One entry point, `workbook`, returning a small mutable builder whose only
// terminal is `bytes()`. Nothing here reads a .xlsx file, and nothing turns a
// string into code: the package runs under a Content Security Policy that
// permits no string-to-code execution, and its suite runs Node with
// `--disallow-code-generation-from-strings`. The source itself is scanned for
// the two constructs, so the words naming them appear nowhere below.
//
// Two promises shape almost every decision below. The same calls produce the
// same bytes on the same runtime, so nothing consults a clock, a locale or a
// random source, and every table is written in a fixed order. And every
// element is emitted in the Open XML SDK's own child sequence, which costs
// nothing here — the parts are assembled as strings in one pass — and removes
// the ordering risk in readers whose tolerance nobody has measured.

// ---------------------------------------------------------------- ZIP ----

// PKWARE APPNOTE 6.3.4, the subset a .xlsx needs: local headers, a central
// directory, an end-of-central-directory record. No ZIP64, no data
// descriptors, no per-entry timestamps.
const CRC = new Int32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  CRC[n] = c;
}

/** @param {Uint8Array} bytes */
const crc32 = (bytes) => {
  let c = -1;
  for (const b of bytes) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
};

/** @param {string} text */
const utf8 = (text) => new TextEncoder().encode(text);

/** @param {Uint8Array} bytes */
const deflate = async (bytes) => {
  const blob = new Blob([/** @type {BlobPart} */ (bytes)]);
  const stream = blob.stream().pipeThrough(new CompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};

// The ZIP epoch, 1980-01-01 00:00. A real timestamp is the one field in the
// container that would vary per run, so it is pinned rather than read.
const DOS_TIME = 0;
const DOS_DATE = (1 << 5) | 1;

// Media is stored rather than deflated: a PNG or JPEG is already compressed,
// so deflating it spends time to grow the file, and method 0 keeps the bytes
// a caller handed over byte-identical inside the package.
const STORED = 0;
const DEFLATED = 8;

/**
 * @param {ReadonlyArray<{ name: string, body: Uint8Array, store?: boolean }>} entries
 * @returns {Promise<Uint8Array>}
 */
const zip = async (entries) => {
  const local = [];
  const central = [];
  let offset = 0;

  for (const entry of entries) {
    const raw = entry.body;
    const packed = entry.store ? raw : await deflate(raw);
    const method = entry.store ? STORED : DEFLATED;
    const name = utf8(entry.name);
    const sum = crc32(raw);

    const head = new DataView(new ArrayBuffer(30));
    head.setUint32(0, 0x04034b50, true);
    head.setUint16(4, 20, true);
    head.setUint16(6, 0x0800, true); // the name is UTF-8
    head.setUint16(8, method, true);
    head.setUint16(10, DOS_TIME, true);
    head.setUint16(12, DOS_DATE, true);
    head.setUint32(14, sum, true);
    head.setUint32(18, packed.length, true);
    head.setUint32(22, raw.length, true);
    head.setUint16(26, name.length, true);
    local.push(new Uint8Array(head.buffer), name, packed);

    const record = new DataView(new ArrayBuffer(46));
    record.setUint32(0, 0x02014b50, true);
    record.setUint16(4, 20, true);
    record.setUint16(6, 20, true);
    record.setUint16(8, 0x0800, true);
    record.setUint16(10, method, true);
    record.setUint16(12, DOS_TIME, true);
    record.setUint16(14, DOS_DATE, true);
    record.setUint32(16, sum, true);
    record.setUint32(20, packed.length, true);
    record.setUint32(24, raw.length, true);
    record.setUint16(28, name.length, true);
    record.setUint32(42, offset, true);
    central.push(new Uint8Array(record.buffer), name);

    offset += 30 + name.length + packed.length;
  }

  const size = central.reduce((total, chunk) => total + chunk.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, size, true);
  end.setUint32(16, offset, true);

  const all = [...local, ...central, new Uint8Array(end.buffer)];
  const bytes = new Uint8Array(all.reduce((total, chunk) => total + chunk.length, 0));
  let at = 0;
  for (const chunk of all) {
    bytes.set(chunk, at);
    at += chunk.length;
  }
  return bytes;
};

// ---------------------------------------------------------------- XML ----

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument";

// What XML 1.0 cannot carry at all: the C0 range except TAB, LF and CR, and
// unpaired surrogates. These are stripped rather than thrown over. The text
// reaching this writer has already passed its caller's own gates, and failing
// a whole document over one stray byte in a customer's name is the wrong
// trade — this is the one place leniency belongs.
const FORBIDDEN =
  // This expression exists to find exactly the control characters the rule
  // below warns about, which is the one case where it is wrong.
  // oxlint-disable-next-line no-control-regex
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

// Excel decodes `_xHHHH_` in cell text back into the character it names, so a
// literal one in author text is escaped at its underscore to survive.
const ESCAPE_SEQUENCE = /_(x[0-9A-Fa-f]{4}_)/g;

/** @param {string} text */
const clean = (text) => text.replace(FORBIDDEN, "").replace(ESCAPE_SEQUENCE, "_x005F_$1");

/** @param {string} text */
const esc = (text) =>
  clean(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

// `xml:space="preserve"` goes on every `<t>` rather than only where the text
// needs it: without it a reader is free to trim, and one rule is easier to
// trust than a predicate over whitespace.
/** @param {string} text */
const textNode = (text) => `<t xml:space="preserve">${esc(text)}</t>`;

/**
 * One run of rich text, written in full: an OOXML run inherits nothing from
 * the cell font.
 * @param {unknown} run
 * @param {string} at
 */
const richRun = (run, at) => {
  if (run === null || typeof run !== "object")
    throw TypeError(`${at}: expected a { text, font } run, got ${JSON.stringify(run)}`);
  const text = /** @type {Record<string, unknown>} */ (run).text;
  if (typeof text !== "string")
    throw TypeError(`${at}.text: expected a string, got ${JSON.stringify(text)}`);
  const font = /** @type {Record<string, unknown>} */ (run).font;
  const properties =
    font === undefined || font === null
      ? ""
      : `<rPr>${runFontXml(/** @type {Record<string, unknown>} */ (font), `${at}.font`)}</rPr>`;
  return `<r>${properties}${textNode(text)}</r>`;
};

/**
 * One `si` of rich text.
 * @param {ReadonlyArray<unknown>} runs
 * @param {string} where
 */
const richXml = (runs, where) =>
  "<si>" + runs.map((run, index) => richRun(run, `${where}[${index}]`)).join("") + "</si>";

/**
 * The interning key for a run list. Prefixed so a one-run list can never
 * collide with the plain string it would render as.
 * @param {ReadonlyArray<unknown>} runs
 * @param {string} where
 */
const richKey = (runs, where) => {
  if (runs.length === 0)
    throw RangeError(`${where}: expected at least one run, got an empty array`);
  return "r" + JSON.stringify(runs);
};

// ------------------------------------------------------------- values ----

// Excel's 1900 date system counts from an epoch of 1899-12-30, which absorbs
// its phantom 1900-02-29. The conversion is in UTC: a reader's own timezone
// is never consulted, so one `Date` is one serial everywhere.
//
// The phantom day is real for serial 60 and nothing else. Excel shows 60 as
// 1900-02-29, a day that never happened, so every date before 1900-03-01 sits
// one lower than the 1899-12-30 epoch alone would put it. A writer that skips
// this arithmetic moves 1900-02-28 onto the phantom day.
const EPOCH_OFFSET = 25569;
const DAY = 86_400_000;
const PHANTOM = 61;

/** @param {Date} date */
const serial = (date) => {
  const days = EPOCH_OFFSET + date.getTime() / DAY;
  return days < PHANTOM ? days - 1 : days;
};

// Excel's built-in "short date". A `Date` written with no format of its own
// reads back as a bare serial — 45351.5 rather than a day — so the writer
// supplies this one, and a reader renders it in its own locale.
const SHORT_DATE = "mm-dd-yy";

// JavaScript writes an exponent as `e+21` and Excel writes `E+21`; readers
// are not uniformly happy with the lowercase form.
/** @param {number} value */
const number = (value) => String(value).replace("e", "E");

/**
 * A 1-based column index as its letters: 1 is A, 27 is AA, 703 is AAA.
 * @param {number} index
 */
const letters = (index) => {
  let out = "";
  let rest = index;
  while (rest > 0) {
    const remainder = (rest - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    rest = (rest - remainder - 1) / 26;
  }
  return out;
};

const COLOUR = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;

/**
 * `#rrggbb` (or `#rgb`) as the `AARRGGBB` OOXML spells. There is no alpha in
 * this surface: a colour either paints or is absent.
 * @param {unknown} value
 * @param {string} where
 */
const rgb = (value, where) => {
  if (typeof value !== "string" || !COLOUR.test(value))
    throw TypeError(`${where}: expected a #rrggbb colour, got ${JSON.stringify(value)}`);
  const body = value.slice(1);
  const full = body.length === 3 ? body[0] + body[0] + body[1] + body[1] + body[2] + body[2] : body;
  return "FF" + full.toUpperCase();
};

// The number formats every reader knows by id. A caller naming one of these
// codes gets the id; anything else is interned from 164 up.
const BUILT_IN = new Map([
  ["General", 0],
  ["0", 1],
  ["0.00", 2],
  ["#,##0", 3],
  ["#,##0.00", 4],
  ["0%", 9],
  ["0.00%", 10],
  ["0.00E+00", 11],
  ["# ?/?", 12],
  ["# ??/??", 13],
  ["mm-dd-yy", 14],
  ["d-mmm-yy", 15],
  ["d-mmm", 16],
  ["mmm-yy", 17],
  ["h:mm AM/PM", 18],
  ["h:mm:ss AM/PM", 19],
  ["h:mm", 20],
  ["h:mm:ss", 21],
  ["m/d/yy h:mm", 22],
  ["mm:ss", 45],
  ["[h]:mm:ss", 46],
  ["mmss.0", 47],
  ["##0.0E+0", 48],
  ["@", 49],
]);

const FIRST_CUSTOM = 164;

// ------------------------------------------------------------- tables ----

/**
 * Interning by a canonical key, preserving first-seen order. Every style
 * table in the file is one of these, which is what makes identical fonts,
 * fills, borders and formats collapse to one entry each.
 */
/**
 * @template T
 * @param {Array<[string, T]>} initial
 */
const table = (initial) => {
  /** @type {Map<string, number>} */
  const keys = new Map();
  /** @type {T[]} */
  const items = [];
  for (const [key, item] of initial) {
    keys.set(key, items.length);
    items.push(item);
  }
  return {
    items,
    /**
     * Arrow rather than a method, so `strings` below can pass it on by name:
     * it closes over the table and never touches `this`.
     * @param {string} key
     * @param {() => T} make
     */
    intern: (key, make) => {
      const seen = keys.get(key);
      if (seen !== undefined) return seen;
      const at = items.length;
      keys.set(key, at);
      items.push(make());
      return at;
    },
  };
};

// ------------------------------------------------------------- styles ----

/**
 * The font half of a style, written in the SDK's `CT_Font` sequence:
 * b i strike u sz color name.
 * @param {Record<string, unknown>} font
 * @param {string} where
 */
const fontXml = (font, where) => {
  let out = "";
  if (font.bold) out += "<b/>";
  if (font.italic) out += "<i/>";
  if (font.strikethrough) out += "<strike/>";
  if (font.underline) out += "<u/>";
  out += `<sz val="${font.size === undefined ? 11 : requireNumber(font.size, `${where}.size`)}"/>`;
  if (font.color !== undefined) out += `<color rgb="${rgb(font.color, `${where}.color`)}"/>`;
  return out + `<name val="${esc(requireText(font.name, `${where}.name`, "Calibri"))}"/>`;
};

/**
 * A run's font, identical to a cell font but for the tag names OOXML uses
 * inside `rPr`, where `name` is spelled `rFont`.
 * @param {Record<string, unknown>} font
 * @param {string} where
 */
const runFontXml = (font, where) => fontXml(font, where).replace(/<name val=/, "<rFont val=");

const BORDER_STYLES = new Set(["thin", "dashed", "dotted"]);
const SIDES = ["left", "right", "top", "bottom"];

/**
 * @param {Record<string, unknown>} border
 * @param {string} where
 */
const borderXml = (border, where) => {
  let out = "";
  for (const side of SIDES) {
    const edge = border[side];
    if (edge === undefined || edge === null) {
      out += `<${side}/>`;
      continue;
    }
    const at = `${where}.${side}`;
    const style = /** @type {Record<string, unknown>} */ (edge).style;
    if (typeof style !== "string" || !BORDER_STYLES.has(style))
      throw RangeError(
        `${at}.style: expected thin, dashed or dotted, got ${JSON.stringify(style)}`,
      );
    const colour = /** @type {Record<string, unknown>} */ (edge).color;
    out +=
      colour === undefined
        ? `<${side} style="${style}"/>`
        : `<${side} style="${style}"><color rgb="${rgb(colour, `${at}.color`)}"/></${side}>`;
  }
  return out + "<diagonal/>";
};

const HORIZONTAL = new Set(["left", "center", "right", "justify", "fill"]);
const VERTICAL = new Set(["top", "center", "bottom", "justify"]);

/**
 * @param {Record<string, unknown>} alignment
 * @param {string} where
 */
const alignmentXml = (alignment, where) => {
  let out = "";
  if (alignment.horizontal !== undefined) {
    const horizontal = /** @type {string} */ (alignment.horizontal);
    if (!HORIZONTAL.has(horizontal))
      throw RangeError(`${where}.horizontal: unknown alignment ${JSON.stringify(horizontal)}`);
    out += ` horizontal="${horizontal}"`;
  }
  if (alignment.vertical !== undefined) {
    const vertical = /** @type {string} */ (alignment.vertical);
    if (!VERTICAL.has(vertical))
      throw RangeError(`${where}.vertical: unknown alignment ${JSON.stringify(vertical)}`);
    out += ` vertical="${vertical}"`;
  }
  if (alignment.wrapText) out += ' wrapText="1"';
  return out;
};

/**
 * @param {unknown} value
 * @param {string} where
 * @param {string} fallback
 */
const requireText = (value, where, fallback) => {
  if (value === undefined) return fallback;
  if (typeof value !== "string")
    throw TypeError(`${where}: expected a string, got ${JSON.stringify(value)}`);
  return value;
};

/**
 * @param {unknown} value
 * @param {string} where
 */
const requireNumber = (value, where) => {
  if (typeof value !== "number" || !Number.isFinite(value))
    throw RangeError(`${where}: expected a finite number, got ${JSON.stringify(value)}`);
  return value;
};

/**
 * @typedef {{ numFmtId: number, fontId: number, fillId: number, borderId: number, alignment: string }} Xf
 */

/**
 * One `xf`. Each `apply*` flag says which table this format actually reaches
 * into; a reader ignores a table entry the flag does not claim.
 * @param {Xf} xf
 */
const xfXml = (xf) => {
  let out = `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"`;
  if (xf.numFmtId !== 0) out += ' applyNumberFormat="1"';
  if (xf.fontId !== 0) out += ' applyFont="1"';
  if (xf.fillId !== 0) out += ' applyFill="1"';
  if (xf.borderId !== 0) out += ' applyBorder="1"';
  if (xf.alignment !== "") out += ' applyAlignment="1"';
  return xf.alignment === "" ? out + "/>" : `${out}><alignment${xf.alignment}/></xf>`;
};

/**
 * The whole of `styles.xml`, and the interning behind it. Entry 0 of each
 * table is the default every unstyled cell points at; the two fills OOXML
 * reserves (`none`, `gray125`) are written whether or not anything uses them,
 * because readers index fills by position.
 */
const stylesheet = () => {
  const numFmts = new Map();
  const fonts = table([["", `<font><sz val="11"/><name val="Calibri"/></font>`]]);
  const fills = table([
    ["none", `<fill><patternFill patternType="none"/></fill>`],
    ["gray125", `<fill><patternFill patternType="gray125"/></fill>`],
  ]);
  const borders = table([["", `<border><left/><right/><top/><bottom/><diagonal/></border>`]]);
  const formats = table([["", { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: "" }]]);

  /**
   * A style part that is absent means entry 0 of its table — the default every
   * unstyled cell already points at.
   *
   * @template T
   * @param {unknown} part
   * @param {{ intern(key: string, make: () => T): number }} into
   * @param {(value: Record<string, unknown>) => [string, () => T]} describe
   */
  const optional = (part, into, describe) => {
    if (part === undefined || part === null) return 0;
    const [key, make] = describe(/** @type {Record<string, unknown>} */ (part));
    return into.intern(key, make);
  };

  /**
   * A format code a reader knows by id keeps that id; anything else is
   * interned from 164 up, in first-seen order.
   * @param {unknown} code
   * @param {string} where
   */
  const formatId = (code, where) => {
    if (code === undefined || code === null) return 0;
    if (typeof code !== "string")
      throw TypeError(`${where}: expected a format code, got ${JSON.stringify(code)}`);
    const builtIn = BUILT_IN.get(code);
    if (builtIn !== undefined) return builtIn;
    const seen = numFmts.get(code);
    if (seen !== undefined) return seen;
    const id = FIRST_CUSTOM + numFmts.size;
    numFmts.set(code, id);
    return id;
  };

  return {
    /**
     * @param {Record<string, unknown> | null | undefined} style
     * @param {string} where
     * @returns {number}
     */
    intern(style, where) {
      if (style === undefined || style === null) return 0;
      if (typeof style !== "object")
        throw TypeError(`${where}: expected a style object, got ${JSON.stringify(style)}`);

      const record = {
        numFmtId: formatId(style.numberFormat, `${where}.numberFormat`),
        fontId: optional(style.font, fonts, (font) => [
          JSON.stringify(font),
          () => `<font>${fontXml(font, `${where}.font`)}</font>`,
        ]),
        fillId: optional(style.fill, fills, () => {
          const colour = rgb(style.fill, `${where}.fill`);
          return [
            colour,
            () =>
              `<fill><patternFill patternType="solid"><fgColor rgb="${colour}"/><bgColor indexed="64"/></patternFill></fill>`,
          ];
        }),
        borderId: optional(style.border, borders, (border) => [
          JSON.stringify(border),
          () => `<border>${borderXml(border, `${where}.border`)}</border>`,
        ]),
        alignment:
          style.alignment === undefined || style.alignment === null
            ? ""
            : alignmentXml(
                /** @type {Record<string, unknown>} */ (style.alignment),
                `${where}.alignment`,
              ),
      };
      return formats.intern(JSON.stringify(record), () => record);
    },

    xml() {
      const custom = [...numFmts]
        .map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`)
        .join("");
      const cellXfs = formats.items.map(xfXml).join("");

      return (
        DECLARATION +
        `<styleSheet xmlns="${NS}">` +
        (custom === "" ? "" : `<numFmts count="${numFmts.size}">${custom}</numFmts>`) +
        `<fonts count="${fonts.items.length}">${fonts.items.join("")}</fonts>` +
        `<fills count="${fills.items.length}">${fills.items.join("")}</fills>` +
        `<borders count="${borders.items.length}">${borders.items.join("")}</borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="${formats.items.length}">${cellXfs}</cellXfs>` +
        `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
        `</styleSheet>`
      );
    },
  };
};

// ------------------------------------------------------ shared strings ----

/**
 * Text is always interned here rather than written inline. It is what Excel
 * itself writes, and a report's repeated labels are what it shrinks.
 */
const strings = () => {
  // The same interning the style tables use: one `si` per distinct piece of
  // text, in first-seen order.
  const entries = table(/** @type {Array<[string, string]>} */ ([]));
  return {
    intern: entries.intern,
    xml() {
      const count = entries.items.length;
      return (
        DECLARATION +
        `<sst xmlns="${NS}" count="${count}" uniqueCount="${count}">` +
        entries.items.join("") +
        `</sst>`
      );
    },
  };
};

// -------------------------------------------------------------- sheet ----

const FORBIDDEN_IN_NAME = /[:\\/?*[\]]/;
const NAME_LIMIT = 31;

// A sheet is at most 1,048,576 rows by 16,384 columns. The writer refuses
// past that rather than emitting a file a reader silently truncates.
const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;

/**
 * @param {unknown} value
 * @param {number} limit
 * @param {string} where
 */
const requireIndex = (value, limit, where) => {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1 || value > limit)
    throw RangeError(
      `${where}: expected an integer between 1 and ${limit}, got ${JSON.stringify(value)}`,
    );
  return value;
};

// EMU per CSS pixel at 96 dpi.
const EMU = 9525;

/**
 * One worksheet. Rows accumulate in call order; everything else is bookkeeping
 * the sheet part needs at the end.
 *
 * @param {string} name
 * @param {ReturnType<typeof stylesheet>} styles
 * @param {ReturnType<typeof strings>} sst
 * @param {(id: unknown) => boolean} knows
 */
const worksheet = (name, styles, sst, knows) => {
  /** @type {string[]} */
  const rows = [];
  /** @type {Array<{ top: number, left: number, bottom: number, right: number }>} */
  const merges = [];
  /** @type {Array<{ id: number, row: number, col: number, width: number, height: number }>} */
  const pictures = [];
  let frozen = 0;
  let widest = 1;

  /**
   * @param {unknown} value
   * @param {number} styleId
   * @param {string} ref
   * @param {string} where
   */
  const cellXml = (value, styleId, ref, where) => {
    const s = styleId === 0 ? "" : ` s="${styleId}"`;

    if (value === undefined || value === null) return s === "" ? "" : `<c r="${ref}"${s}/>`;

    if (typeof value === "number") {
      requireNumber(value, `${where}.value`);
      return `<c r="${ref}"${s}><v>${number(value)}</v></c>`;
    }

    if (typeof value === "boolean") return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;

    if (value instanceof Date) {
      if (Number.isNaN(value.getTime()))
        throw RangeError(`${where}.value: expected a valid Date, got an invalid one`);
      return `<c r="${ref}"${s}><v>${number(serial(value))}</v></c>`;
    }

    if (typeof value === "string") {
      const at = sst.intern("s" + value, () => `<si>${textNode(value)}</si>`);
      return `<c r="${ref}"${s} t="s"><v>${at}</v></c>`;
    }

    if (Array.isArray(value)) {
      const at = sst.intern(richKey(value, `${where}.value`), () =>
        richXml(value, `${where}.value`),
      );
      return `<c r="${ref}"${s} t="s"><v>${at}</v></c>`;
    }

    throw TypeError(`${where}.value: unsupported cell value ${JSON.stringify(value)}`);
  };

  return {
    name,

    /** @param {ReadonlyArray<unknown>} cells */
    row(cells) {
      if (!Array.isArray(cells))
        throw TypeError(`row: expected an array of cells, got ${JSON.stringify(cells)}`);
      const at = rows.length + 1;
      if (cells.length > MAX_COLUMN)
        throw RangeError(`row: a sheet holds at most ${MAX_COLUMN} columns`);

      let body = "";
      for (let index = 0; index < cells.length; index++) {
        const cell = cells[index];
        if (cell === undefined || cell === null) continue;
        if (typeof cell !== "object")
          throw TypeError(`row: cell ${index + 1} is not a { value, style } object`);
        const where = `row ${at}, cell ${index + 1}`;
        const declared = /** @type {Record<string, unknown> | null | undefined} */ (
          /** @type {Record<string, unknown>} */ (cell).style
        );
        // A date with no format of its own gets the short-date built-in, so a
        // `Date` reads back as a day rather than as the number underneath it.
        // A style that is not an object passes through untouched, so the
        // interning below is still the one place that refuses it.
        const usable = declared === null || declared === undefined || typeof declared === "object";
        const style =
          /** @type {Record<string, unknown>} */ (cell).value instanceof Date &&
          usable &&
          (declared?.numberFormat ?? null) === null
            ? { ...declared, numberFormat: SHORT_DATE }
            : /** @type {Record<string, unknown> | null | undefined} */ (declared);
        const styleId = styles.intern(style, `${where}.style`);
        body += cellXml(
          /** @type {Record<string, unknown>} */ (cell).value,
          styleId,
          letters(index + 1) + at,
          where,
        );
      }
      if (cells.length > widest) widest = cells.length;
      rows.push(`<row r="${at}">${body}</row>`);
      return at;
    },

    /**
     * @param {number} row
     * @param {number} at
     * @param {number} width
     */
    merge(row, at, width) {
      requireIndex(row, MAX_ROW, "merge: row");
      requireIndex(at, MAX_COLUMN, "merge: at");
      if (row > rows.length)
        throw RangeError(`merge: row ${row} does not exist yet (the sheet has ${rows.length})`);
      if (typeof width !== "number" || !Number.isInteger(width) || width < 2)
        throw RangeError(`merge: expected a width of 2 or more, got ${JSON.stringify(width)}`);
      const right = at + width - 1;
      if (right > MAX_COLUMN) throw RangeError(`merge: the range ends past column ${MAX_COLUMN}`);
      for (const other of merges)
        if (row >= other.top && row <= other.bottom && at <= other.right && right >= other.left)
          throw RangeError(
            `merge: the range ${letters(at)}${row}:${letters(right)}${row} overlaps ` +
              `${letters(other.left)}${other.top}:${letters(other.right)}${other.bottom}`,
          );
      merges.push({ top: row, left: at, bottom: row, right });
      if (right > widest) widest = right;
    },

    /** @param {number} count */
    freeze(count) {
      if (typeof count !== "number" || !Number.isInteger(count) || count < 0 || count >= MAX_ROW)
        throw RangeError(`freeze: expected a row count of 0 or more, got ${JSON.stringify(count)}`);
      frozen = count;
    },

    /**
     * @param {unknown} id
     * @param {{ row: number, col?: number, width: number, height: number }} at
     */
    place(id, at) {
      if (!knows(id)) throw TypeError(`place: no image with id ${JSON.stringify(id)}`);
      if (at === null || typeof at !== "object")
        throw TypeError(`place: expected { row, col, width, height }, got ${JSON.stringify(at)}`);
      const row = requireIndex(at.row, MAX_ROW, "place: row");
      const col = requireIndex(at.col === undefined ? 1 : at.col, MAX_COLUMN, "place: col");
      const width = requireNumber(at.width, "place: width");
      const height = requireNumber(at.height, "place: height");
      if (width <= 0 || height <= 0)
        throw RangeError(`place: expected a positive size, got ${width} by ${height}`);
      pictures.push({ id: /** @type {number} */ (id), row, col, width, height });
    },

    get pictures() {
      return pictures;
    },

    /** @param {number | null} drawing the relationship id of this sheet's drawing part */
    xml(drawing) {
      const dimension = `A1:${letters(widest)}${Math.max(rows.length, 1)}`;
      const pane =
        frozen === 0
          ? '<sheetView workbookViewId="0"/>'
          : `<sheetView workbookViewId="0"><pane ySplit="${frozen}" topLeftCell="A${frozen + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${frozen + 1}" sqref="A${frozen + 1}"/></sheetView>`;
      const merged =
        merges.length === 0
          ? ""
          : `<mergeCells count="${merges.length}">` +
            merges
              .map(
                (range) =>
                  `<mergeCell ref="${letters(range.left)}${range.top}:${letters(range.right)}${range.bottom}"/>`,
              )
              .join("") +
            `</mergeCells>`;

      return (
        DECLARATION +
        `<worksheet xmlns="${NS}" xmlns:r="${NS_R}">` +
        `<dimension ref="${dimension}"/>` +
        `<sheetViews>${pane}</sheetViews>` +
        `<sheetData>${rows.join("")}</sheetData>` +
        merged +
        (drawing === null ? "" : `<drawing r:id="rId${drawing}"/>`) +
        `</worksheet>`
      );
    },
  };
};

/**
 * The drawing part for one sheet: a `oneCellAnchor` per placement, anchored to
 * the top-left of its cell at the size the caller asked for. The XML counts
 * rows and columns from zero and the subtraction happens here, so the surface
 * stays 1-based throughout.
 *
 * @param {ReturnType<typeof worksheet>["pictures"]} pictures
 * @param {ReadonlyArray<number>} media the media index each picture points at
 */
const drawingXml = (pictures, media) =>
  DECLARATION +
  `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  pictures
    .map((picture, index) => {
      const cx = Math.round(picture.width * EMU);
      const cy = Math.round(picture.height * EMU);
      return (
        `<xdr:oneCellAnchor editAs="oneCell">` +
        `<xdr:from><xdr:col>${picture.col - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${picture.row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:ext cx="${cx}" cy="${cy}"/>` +
        `<xdr:pic>` +
        `<xdr:nvPicPr><xdr:cNvPr id="${index + 2}" name="Picture ${index + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip xmlns:r="${NS_R}" r:embed="rId${media[index] + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
        `</xdr:pic>` +
        `<xdr:clientData/>` +
        `</xdr:oneCellAnchor>`
      );
    })
    .join("") +
  `</xdr:wsDr>`;

// ----------------------------------------------------------- workbook ----

const FORMATS = new Map([
  ["png", "image/png"],
  ["jpeg", "image/jpeg"],
]);

// Pinned rather than stamped. Byte-identity for the same input is this
// writer's headline promise, and an overridable clock would quietly redefine
// "the same input" as "the same input at the same wall time".
const CREATED = "1970-01-01T00:00:00Z";

/** @param {Uint8Array} a @param {Uint8Array} b */
const same = (a, b) => {
  if (a.length !== b.length) return false;
  for (let at = 0; at < a.length; at++) if (a[at] !== b[at]) return false;
  return true;
};

/**
 * Open a workbook.
 *
 * @param {{ title?: string, creator?: string, subject?: string, description?: string }} [meta]
 */
export const workbook = (meta = {}) => {
  if (meta === null || typeof meta !== "object")
    throw TypeError(`workbook: expected a metadata object, got ${JSON.stringify(meta)}`);

  const styles = stylesheet();
  const sst = strings();
  /** @type {Array<{ bytes: Uint8Array, format: string }>} */
  const media = [];
  /** @type {Array<ReturnType<typeof worksheet>>} */
  const sheets = [];

  /** @param {"title" | "creator" | "subject" | "description"} key */
  const property = (key) => {
    const value = meta[key];
    if (value === undefined) return undefined;
    if (typeof value !== "string")
      throw TypeError(`workbook: ${key} must be a string, got ${JSON.stringify(value)}`);
    return value;
  };

  const title = property("title");
  const creator = property("creator");
  const subject = property("subject");
  const description = property("description");

  return {
    /**
     * @param {Uint8Array} bytes
     * @param {"png" | "jpeg"} format
     */
    image(bytes, format) {
      if (!(bytes instanceof Uint8Array))
        throw TypeError(`image: expected a Uint8Array, got ${JSON.stringify(bytes)}`);
      if (!FORMATS.has(format))
        throw TypeError(`image: expected png or jpeg, got ${JSON.stringify(format)}`);
      // Identical bytes deduplicate: a caller embedding the same logo on every
      // sheet pays for it once, and does so without keeping a tally of its own.
      for (let at = 0; at < media.length; at++)
        if (media[at].format === format && same(media[at].bytes, bytes)) return at;
      media.push({ bytes, format });
      return media.length - 1;
    },

    /** @param {string} name */
    sheet(name) {
      if (typeof name !== "string" || name === "")
        throw RangeError(`sheet: expected a name, got ${JSON.stringify(name)}`);
      if (name.length > NAME_LIMIT)
        throw RangeError(`sheet: a name is at most ${NAME_LIMIT} characters, got ${name.length}`);
      if (FORBIDDEN_IN_NAME.test(name))
        throw RangeError(
          `sheet: a name cannot contain : \\ / ? * [ ], got ${JSON.stringify(name)}`,
        );
      if (sheets.some((each) => each.name === name))
        throw RangeError(`sheet: a sheet named ${JSON.stringify(name)} already exists`);
      const made = worksheet(
        name,
        styles,
        sst,
        (id) => typeof id === "number" && Number.isInteger(id) && id >= 0 && id < media.length,
      );
      sheets.push(made);
      return made;
    },

    async bytes() {
      if (sheets.length === 0) throw RangeError("bytes: a workbook needs at least one sheet");

      /** @type {Array<{ name: string, body: Uint8Array, store?: boolean }>} */
      const entries = [];
      /**
       * @param {string} name
       * @param {string | Uint8Array} body
       * @param {boolean} [store]
       */
      const part = (name, body, store) =>
        entries.push({
          name,
          body: store ? /** @type {Uint8Array} */ (body) : utf8(/** @type {string} */ (body)),
          store,
        });

      // Which media each sheet's drawing references, and in what order, so a
      // drawing's relationship ids line up with its own part rather than with
      // the workbook-wide media list.
      const drawings = sheets
        .map((sheet, index) => ({ sheet, index }))
        .filter((each) => each.sheet.pictures.length > 0);

      const overrides = sheets
        .map(
          (_sheet, index) =>
            `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="${CT}.spreadsheetml.worksheet+xml"/>`,
        )
        .join("");
      const drawingOverrides = drawings
        .map(
          (_each, at) =>
            `<Override PartName="/xl/drawings/drawing${at + 1}.xml" ContentType="${CT}.drawing+xml"/>`,
        )
        .join("");
      const defaults = [...new Set(media.map((each) => each.format))]
        .map((format) => `<Default Extension="${format}" ContentType="${FORMATS.get(format)}"/>`)
        .join("");

      part(
        "[Content_Types].xml",
        DECLARATION +
          `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          defaults +
          `<Override PartName="/xl/workbook.xml" ContentType="${CT}.spreadsheetml.sheet.main+xml"/>` +
          overrides +
          `<Override PartName="/xl/styles.xml" ContentType="${CT}.spreadsheetml.styles+xml"/>` +
          `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}.spreadsheetml.sharedStrings+xml"/>` +
          drawingOverrides +
          `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
          `<Override PartName="/docProps/app.xml" ContentType="${CT}.extended-properties+xml"/>` +
          `</Types>`,
      );

      part(
        "_rels/.rels",
        DECLARATION +
          `<Relationships xmlns="${NS_PKG_REL}">` +
          `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>` +
          `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
          `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
          `</Relationships>`,
      );

      /** @param {string} name @param {string | undefined} value */
      const tag = (name, value) => (value === undefined ? "" : `<${name}>${esc(value)}</${name}>`);
      part(
        "docProps/core.xml",
        DECLARATION +
          `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
          tag("dc:title", title) +
          tag("dc:subject", subject) +
          tag("dc:creator", creator) +
          tag("dc:description", description) +
          `<dcterms:created xsi:type="dcterms:W3CDTF">${CREATED}</dcterms:created>` +
          `<dcterms:modified xsi:type="dcterms:W3CDTF">${CREATED}</dcterms:modified>` +
          `</cp:coreProperties>`,
      );

      part(
        "docProps/app.xml",
        DECLARATION +
          `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>werkmap</Application></Properties>`,
      );

      part(
        "xl/workbook.xml",
        DECLARATION +
          `<workbook xmlns="${NS}" xmlns:r="${NS_R}"><sheets>` +
          sheets
            .map(
              (sheet, index) =>
                `<sheet name="${esc(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
            )
            .join("") +
          `</sheets></workbook>`,
      );

      part(
        "xl/_rels/workbook.xml.rels",
        DECLARATION +
          `<Relationships xmlns="${NS_PKG_REL}">` +
          sheets
            .map(
              (_sheet, index) =>
                `<Relationship Id="rId${index + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
            )
            .join("") +
          `<Relationship Id="rId${sheets.length + 1}" Type="${REL}/styles" Target="styles.xml"/>` +
          `<Relationship Id="rId${sheets.length + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>` +
          `</Relationships>`,
      );

      for (const [index, sheet] of sheets.entries()) {
        const at = drawings.findIndex((each) => each.index === index);
        part(`xl/worksheets/sheet${index + 1}.xml`, sheet.xml(at === -1 ? null : 1));
        if (at !== -1)
          part(
            `xl/worksheets/_rels/sheet${index + 1}.xml.rels`,
            DECLARATION +
              `<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing${at + 1}.xml"/></Relationships>`,
          );
      }

      part("xl/styles.xml", styles.xml());
      part("xl/sharedStrings.xml", sst.xml());

      for (const [at, each] of drawings.entries()) {
        // Each drawing numbers its own relationships from one, in the order
        // its pictures were placed.
        const used = [...new Set(each.sheet.pictures.map((picture) => picture.id))];
        part(
          `xl/drawings/drawing${at + 1}.xml`,
          drawingXml(
            each.sheet.pictures,
            each.sheet.pictures.map((picture) => used.indexOf(picture.id)),
          ),
        );
        part(
          `xl/drawings/_rels/drawing${at + 1}.xml.rels`,
          DECLARATION +
            `<Relationships xmlns="${NS_PKG_REL}">` +
            used
              .map(
                (id, index) =>
                  `<Relationship Id="rId${index + 1}" Type="${REL}/image" Target="../media/image${id + 1}.${media[id].format}"/>`,
              )
              .join("") +
            `</Relationships>`,
        );
      }

      for (const [index, each] of media.entries())
        part(`xl/media/image${index + 1}.${each.format}`, each.bytes, true);

      return zip(entries);
    },
  };
};
