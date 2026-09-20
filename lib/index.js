// werkmap — a write-only OOXML (.xlsx) writer.
//
// One entry point, `workbook`, returning a mutable builder whose only terminal
// is `bytes()`. Nothing here reads a .xlsx file.

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

// Indexed rather than `for...of`: this runs over every part's uncompressed
// bytes, so it scales with the whole document, and the iterator protocol costs
// about four times the arithmetic it wraps (32 ms against 9 ms over 3 MB).
/** @param {Uint8Array} bytes */
const crc32 = (bytes) => {
  let c = -1;
  for (let at = 0; at < bytes.length; at++) c = CRC[(c ^ bytes[at]) & 0xff] ^ (c >>> 8);
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
 * @typedef {{ method: number, sum: number, packed: Uint8Array, raw: Uint8Array, name: Uint8Array }} Packed
 */

/** @param {Packed} entry */
const localHeader = ({ method, sum, packed, raw, name }) => {
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
  return new Uint8Array(head.buffer);
};

/** @param {Packed} entry @param {number} offset */
const centralRecord = ({ method, sum, packed, raw, name }, offset) => {
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
  return new Uint8Array(record.buffer);
};

/** @param {number} count @param {number} size @param {number} offset */
const endRecord = (count, size, offset) => {
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, count, true);
  end.setUint16(10, count, true);
  end.setUint32(12, size, true);
  end.setUint32(16, offset, true);
  return new Uint8Array(end.buffer);
};

/** @param {ReadonlyArray<Uint8Array>} chunks */
const concat = (chunks) => {
  const bytes = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
  let at = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, at);
    at += chunk.length;
  }
  return bytes;
};

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
    const [packed, method] = entry.store ? [raw, STORED] : [await deflate(raw), DEFLATED];
    const name = utf8(entry.name);
    const item = { method, sum: crc32(raw), packed, raw, name };
    local.push(localHeader(item), name, packed);
    central.push(centralRecord(item, offset), name);
    offset += 30 + name.length + packed.length;
  }

  const size = central.reduce((total, chunk) => total + chunk.length, 0);
  return concat([...local, ...central, endRecord(entries.length, size, offset)]);
};

// ---------------------------------------------------------------- XML ----

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument";

// What XML 1.0 cannot carry at all: the C0 range except TAB, LF and CR, and
// unpaired surrogates. Stripped rather than thrown over — the one place
// leniency belongs.
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

// `xml:space="preserve"` goes on every `<t>`, not only where the text needs
// it: one rule is easier to trust than a predicate over whitespace.
/** @param {string} text */
const textNode = (text) => `<t xml:space="preserve">${esc(text)}</t>`;

// ---------------------------------------------------------- validation ----

// Absent means the caller said nothing: `undefined` and `null` read the same
// everywhere on this surface.
/** @param {unknown} value */
const absent = (value) => value === undefined || value === null;

/**
 * @param {unknown} value
 * @param {string} where
 * @param {string} expected what the message names when the value is not an object
 * @returns {Record<string, unknown>}
 */
const requireRecord = (value, where, expected) => {
  if (value === null || typeof value !== "object")
    throw TypeError(`${where}: expected ${expected}, got ${JSON.stringify(value)}`);
  return /** @type {Record<string, unknown>} */ (value);
};

/**
 * @param {unknown} value
 * @param {string} where
 */
const requireString = (value, where) => {
  if (typeof value !== "string")
    throw TypeError(`${where}: expected a string, got ${JSON.stringify(value)}`);
  return value;
};

/**
 * @param {unknown} value
 * @param {string} where
 * @param {string} fallback
 */
const requireText = (value, where, fallback) =>
  value === undefined ? fallback : requireString(value, where);

/**
 * @param {unknown} value
 * @param {string} where
 */
const requireNumber = (value, where) => {
  if (typeof value !== "number" || !Number.isFinite(value))
    throw RangeError(`${where}: expected a finite number, got ${JSON.stringify(value)}`);
  return value;
};

/** @param {unknown} value @param {number} limit */
const isIndex = (value, limit) =>
  typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= limit;

/**
 * @param {unknown} value
 * @param {number} limit
 * @param {string} where
 */
const requireIndex = (value, limit, where) => {
  if (!isIndex(value, limit))
    throw RangeError(
      `${where}: expected an integer between 1 and ${limit}, got ${JSON.stringify(value)}`,
    );
  return /** @type {number} */ (value);
};

/**
 * One run of rich text, written in full: an OOXML run inherits nothing from
 * the cell font.
 * @param {unknown} run
 * @param {string} at
 */
const richRun = (run, at) => {
  const { text, font } = requireRecord(run, at, "a { text, font } run");
  const body = textNode(requireString(text, `${at}.text`));
  const properties = absent(font)
    ? ""
    : `<rPr>${runFontXml(/** @type {Record<string, unknown>} */ (font), `${at}.font`)}</rPr>`;
  return `<r>${properties}${body}</r>`;
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

// Excel's 1900 date system counts from an epoch of 1899-12-30, in UTC. The
// phantom 1900-02-29 is serial 60, so every date before 1900-03-01 sits one
// lower than that epoch alone would put it: skip this arithmetic and
// 1900-02-28 lands on the phantom day.
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
 * table is one of these.
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
     * Arrow rather than a method, so `strings` below can pass it on by name.
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

// The boolean half of a font, in the order `CT_Font` writes them.
/** @type {ReadonlyArray<[string, string]>} */
const FONT_FLAGS = [
  ["bold", "<b/>"],
  ["italic", "<i/>"],
  ["strikethrough", "<strike/>"],
  ["underline", "<u/>"],
];

/**
 * The font half of a style, written in the SDK's `CT_Font` sequence:
 * b i strike u sz color name.
 * @param {Record<string, unknown>} font
 * @param {string} where
 */
const fontXml = (font, where) => {
  const flags = FONT_FLAGS.filter(([key]) => font[key])
    .map(([, tag]) => tag)
    .join("");
  const size = font.size === undefined ? 11 : requireNumber(font.size, `${where}.size`);
  const colour =
    font.color === undefined ? "" : `<color rgb="${rgb(font.color, `${where}.color`)}"/>`;
  const name = esc(requireText(font.name, `${where}.name`, "Calibri"));
  return `${flags}<sz val="${size}"/>${colour}<name val="${name}"/>`;
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
/**
 * @param {unknown} style
 * @param {string} at
 */
const edgeStyle = (style, at) => {
  if (typeof style !== "string" || !BORDER_STYLES.has(style))
    throw RangeError(`${at}.style: expected thin, dashed or dotted, got ${JSON.stringify(style)}`);
  return style;
};

/**
 * One side of a border. An absent side is written empty rather than skipped,
 * because `CT_Border` lists all four in a fixed order.
 * @param {string} side
 * @param {unknown} edge
 * @param {string} at
 */
const sideXml = (side, edge, at) => {
  if (absent(edge)) return `<${side}/>`;
  const style = edgeStyle(/** @type {Record<string, unknown>} */ (edge).style, at);
  const colour = /** @type {Record<string, unknown>} */ (edge).color;
  return colour === undefined
    ? `<${side} style="${style}"/>`
    : `<${side} style="${style}"><color rgb="${rgb(colour, `${at}.color`)}"/></${side}>`;
};

/**
 * @param {Record<string, unknown>} border
 * @param {string} where
 */
const borderXml = (border, where) =>
  SIDES.map((side) => sideXml(side, border[side], `${where}.${side}`)).join("") + "<diagonal/>";

const HORIZONTAL = new Set(["left", "center", "right", "justify", "fill"]);
const VERTICAL = new Set(["top", "center", "bottom", "justify"]);

/**
 * One alignment axis as its attribute, or nothing when the caller left it out.
 * @param {unknown} value
 * @param {ReadonlySet<string>} known
 * @param {string} axis
 * @param {string} where
 */
const axisXml = (value, known, axis, where) => {
  if (value === undefined) return "";
  const named = /** @type {string} */ (value);
  if (!known.has(named))
    throw RangeError(`${where}.${axis}: unknown alignment ${JSON.stringify(named)}`);
  return ` ${axis}="${named}"`;
};

/**
 * @param {Record<string, unknown>} alignment
 * @param {string} where
 */
const alignmentXml = (alignment, where) =>
  axisXml(alignment.horizontal, HORIZONTAL, "horizontal", where) +
  axisXml(alignment.vertical, VERTICAL, "vertical", where) +
  (alignment.wrapText ? ' wrapText="1"' : "");

/**
 * @typedef {{ numFmtId: number, fontId: number, fillId: number, borderId: number, alignment: string }} Xf
 */

// Which `apply*` flag each table id switches on, in the order `CT_Xf` lists
// them. Alignment is not a table, so it is written apart.
/** @type {ReadonlyArray<["numFmtId" | "fontId" | "fillId" | "borderId", string]>} */
const APPLIES = [
  ["numFmtId", "applyNumberFormat"],
  ["fontId", "applyFont"],
  ["fillId", "applyFill"],
  ["borderId", "applyBorder"],
];

/**
 * One `xf`. Each `apply*` flag says which table this format actually reaches
 * into; a reader ignores a table entry the flag does not claim.
 * @param {Xf} xf
 */
const xfXml = (xf) => {
  const ids = `numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}"`;
  const applies = APPLIES.filter(([key]) => xf[key] !== 0)
    .map(([, flag]) => ` ${flag}="1"`)
    .join("");
  if (xf.alignment === "") return `<xf ${ids} xfId="0"${applies}/>`;
  return `<xf ${ids} xfId="0"${applies} applyAlignment="1"><alignment${xf.alignment}/></xf>`;
};

/**
 * The whole of `styles.xml`. Entry 0 of each table is the default every
 * unstyled cell points at, and the two fills OOXML reserves (`none`,
 * `gray125`) are always written, because readers index fills by position.
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
    if (absent(part)) return 0;
    const [key, make] = describe(/** @type {Record<string, unknown>} */ (part));
    return into.intern(key, make);
  };

  /** @param {string} code */
  const knownFormat = (code) => BUILT_IN.get(code) ?? numFmts.get(code);

  /**
   * A format code a reader knows by id keeps that id; anything else is
   * interned from 164 up, in first-seen order.
   * @param {unknown} code
   * @param {string} where
   */
  const formatId = (code, where) => {
    if (absent(code)) return 0;
    if (typeof code !== "string")
      throw TypeError(`${where}: expected a format code, got ${JSON.stringify(code)}`);
    const seen = knownFormat(code);
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
      if (absent(style)) return 0;
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
        alignment: absent(style.alignment)
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

/** Text is always interned here rather than written inline. */
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

// --------------------------------------------------------- print setup ----

// The paper sizes OOXML names by number. A workbook carries a code, not a
// width and a height, so this is the whole of what a caller may ask for.
const PAPER = new Map([
  ["letter", 1],
  ["tabloid", 3],
  ["legal", 5],
  ["A3", 8],
  ["A4", 9],
  ["A5", 11],
]);

const ORIENTATION = new Set(["portrait", "landscape"]);

// OOXML states margins in inches; this surface takes points, the unit a print
// margin is written in everywhere else. 72 points to the inch.
const POINTS_PER_INCH = 72;

// The gap Excel leaves for a header and a footer when a caller says nothing.
// `pageMargins` has no optional attributes, so a value is owed either way.
const FURNITURE = 0.3;

// Excel's ceiling for a header or a footer, counted on the string it stores:
// the section code and the doubled ampersands included.
const MAX_FURNITURE = 255;

/**
 * @typedef {string | { field: string } | { text: string, bold?: boolean, size?: number }} PrintPart
 * @typedef {string | ReadonlyArray<PrintPart>} PrintSection
 * @typedef {PrintSection | { left?: PrintSection, center?: PrintSection, right?: PrintSection }} PrintText
 * @typedef {{ margin?: number, size?: string, orientation?: string, fit?: boolean,
 *   titles?: number, header?: PrintText, footer?: PrintText,
 *   firstHeader?: PrintText, firstFooter?: PrintText }} PrintSetup
 */

// The two fields a header may carry, as the format spells them: the page
// number and the page count, which the reader fills in as it paginates.
const FIELDS = new Map([
  ["page", "&P"],
  ["pages", "&N"],
]);

// The three sections of a header, in the order the format reads them.
const SECTIONS = [
  ["left", "&L"],
  ["center", "&C"],
  ["right", "&R"],
];

// The four parts of `headerFooter`, in the schema's order, and the setup key
// each is written from. `differentFirst` rides along once either first-page
// part is present.
const FURNITURE_PARTS = [
  ["header", "oddHeader"],
  ["footer", "oddFooter"],
  ["firstHeader", "firstHeader"],
  ["firstFooter", "firstFooter"],
];

// A header font size is written as two digits, so the format reads exactly
// two and the text that follows is never mistaken for more of the number.
const MAX_HEADER_SIZE = 99;

/**
 * `pageMargins`, `pageSetup` and `headerFooter`, or nothing at all.
 * `pageMargins` has no optional attributes, so any margin means writing all
 * six; the two this surface does not take keep Excel's own header and footer
 * gap.
 *
 * @param {PrintSetup | null} setup
 */
const printXml = (setup) => {
  if (setup === null) return "";
  const attributes = setupAttributes(setup);
  return (
    marginsXml(setup.margin) +
    (attributes === "" ? "" : `<pageSetup${attributes}/>`) +
    furnitureXml(setup)
  );
};

/** Text as the format stores it: an ampersand doubled so none reads as a code.
 * @param {string} text */
const plain = (text) => text.replace(/&/g, "&&");

/**
 * A field as its code.
 * @param {unknown} field
 * @param {string} at
 */
const fieldCode = (field, at) => {
  const code = FIELDS.get(/** @type {string} */ (field));
  if (code === undefined)
    throw RangeError(`${at}: expected a field of page or pages, got ${JSON.stringify(field)}`);
  return code;
};

/**
 * @typedef {{ bold: boolean, size: number | undefined }} Look
 *   What the section's text currently reads as. A code toggles bold and sets a
 *   size until the next code, so a part writes one only where its own look
 *   differs from what stands.
 */

/**
 * @param {unknown} size
 * @param {string} at
 */
const requireHeaderSize = (size, at) => {
  if (!isIndex(size, MAX_HEADER_SIZE))
    throw RangeError(
      `${at}.size: expected an integer between 1 and ${MAX_HEADER_SIZE}, got ${JSON.stringify(size)}`,
    );
  return String(size).padStart(2, "0");
};

/**
 * The bold code a part owes: `&B` toggles, so one is written only where the
 * part's weight differs from what stands.
 * @param {Record<string, unknown>} part
 * @param {Look} look
 */
const boldCode = (part, look) => {
  const bold = part.bold === true;
  const toggles = bold !== look.bold;
  look.bold = bold;
  return toggles ? "&B" : "";
};

/**
 * The size code a part owes: a size holds until the next one, so it is
 * written only where the part names one that differs from what stands.
 * @param {Record<string, unknown>} part
 * @param {string} at
 * @param {Look} look
 */
const sizeCode = (part, at, look) => {
  const size = part.size;
  if (size === undefined || size === look.size) return "";
  look.size = /** @type {number} */ (size);
  return "&" + requireHeaderSize(size, at);
};

/**
 * One part of a section: text, a field as its code, or styled text with the
 * codes that change the look before it.
 * @param {unknown} part
 * @param {string} at
 * @param {Look} look
 */
const furniturePart = (part, at, look) => {
  if (typeof part === "string") return plain(part);
  const record = requireRecord(part, at, "text, a { field } part or a { text } part");
  if (record.field !== undefined) return fieldCode(record.field, at);
  const codes = boldCode(record, look) + sizeCode(record, at, look);
  return codes + plain(requireString(record.text, `${at}.text`));
};

/**
 * One section as the format stores it. A newline is a line break. The look
 * starts plain at every section, since a section is where a code's reach ends.
 * @param {unknown} text
 * @param {string} at
 */
const sectionText = (text, at) => {
  const parts = typeof text === "string" ? [text] : text;
  if (!Array.isArray(parts))
    throw TypeError(`${at}: expected text or an array of parts, got ${JSON.stringify(text)}`);
  /** @type {Look} */
  const look = { bold: false, size: undefined };
  return parts.map((part, index) => furniturePart(part, `${at}[${index}]`, look)).join("");
};

/**
 * A header or footer as the format stores it: one section on the left, or
 * the sections a caller named, each behind its code, in the format's order.
 * @param {unknown} text
 * @param {string} at
 */
const furniture = (text, at) => {
  if (typeof text === "string" || Array.isArray(text)) return "&L" + sectionText(text, at);
  const sections = requireRecord(text, at, "text, parts or { left, center, right } sections");
  return SECTIONS.filter(([key]) => sections[key] !== undefined)
    .map(([key, code]) => code + sectionText(sections[key], `${at}.${key}`))
    .join("");
};

/**
 * `headerFooter`, or nothing for a worksheet that names no part of it. A
 * first-page part makes the first page different, which the element says.
 * @param {PrintSetup} setup
 */
const furnitureXml = (setup) => {
  const record = /** @type {Record<string, unknown>} */ (setup);
  const written = FURNITURE_PARTS.filter(([key]) => record[key] !== undefined).map(
    ([key, tag]) => `<${tag}>${esc(furniture(record[key], tag))}</${tag}>`,
  );
  if (written.length === 0) return "";
  const first = setup.firstHeader !== undefined || setup.firstFooter !== undefined;
  return `<headerFooter${first ? ' differentFirst="1"' : ""}>${written.join("")}</headerFooter>`;
};

/** @param {number | undefined} margin in points */
const marginsXml = (margin) => {
  if (margin === undefined) return "";
  const inches = (margin / POINTS_PER_INCH).toFixed(3);
  return (
    `<pageMargins left="${inches}" right="${inches}" top="${inches}" bottom="${inches}"` +
    ` header="${FURNITURE}" footer="${FURNITURE}"/>`
  );
};

/** @param {{ size?: string, orientation?: string, fit?: boolean }} setup */
const setupAttributes = (setup) => {
  let attributes = "";
  if (setup.size !== undefined) attributes += ` paperSize="${PAPER.get(setup.size)}"`;
  if (setup.orientation !== undefined) attributes += ` orientation="${setup.orientation}"`;
  // `fitToHeight="0"` is what makes it *width* the document fits to: one page
  // across, as many down as it takes.
  if (setup.fit === true) attributes += ' fitToWidth="1" fitToHeight="0"';
  return attributes;
};

// -------------------------------------------------------------- sheet ----

const FORBIDDEN_IN_NAME = /[:\\/?*[\]]/;
const NAME_LIMIT = 31;

// A sheet is at most 1,048,576 rows by 16,384 columns. The writer refuses
// past that rather than emitting a file a reader silently truncates.
const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;

// Excel's ceiling for a column width, in the unit the format states it: a
// count of characters of the default font, the number Excel's own width box
// shows. Not pixels: that conversion runs through the default font's maximum
// digit width, and any cell here may name a font of its own.
const MAX_WIDTH = 255;

// EMU per CSS pixel at 96 dpi.
const EMU = 9525;

// Excel's ceiling for a row outline level.
const MAX_LEVEL = 7;

/** @param {unknown} level */
const isLevel = (level) => level === 0 || isIndex(level, MAX_LEVEL);

// Everything a row may say about itself. Named here because the reader below
// refuses anything else: a key it quietly dropped would be a caller asking for
// a row attribute and getting no error and no attribute.
const ROW_OPTIONS = ["level", "hidden", "collapsed"];

/** @param {unknown} value @param {string} where */
const requireFlag = (value, where) => {
  if (value !== undefined && value !== null && typeof value !== "boolean")
    throw TypeError(`${where}: expected a boolean, got ${JSON.stringify(value)}`);
  return value === true;
};

/** The options object, with nothing in it this writer cannot write. */
/** @param {unknown} options */
const requireOptions = (options) => {
  const said = requireRecord(options ?? {}, "row: options", `{ ${ROW_OPTIONS.join(", ")} }`);
  const stray = Object.keys(said).find((name) => !ROW_OPTIONS.includes(name));
  if (stray !== undefined) throw TypeError(`row: options: unknown option ${JSON.stringify(stray)}`);
  return said;
};

/** @param {unknown} level */
const requireOutline = (level) => {
  if (!isLevel(level))
    throw RangeError(
      `row: level: expected an integer between 0 and ${MAX_LEVEL}, got ${JSON.stringify(level)}`,
    );
  return /** @type {number} */ (level);
};

/**
 * What a row says about itself: its outline level, and whether it is hidden or
 * carries a collapsed group's control.
 * @param {unknown} options
 */
const requireRow = (options) => {
  const said = requireOptions(options);
  return {
    level: requireOutline(said.level ?? 0),
    hidden: requireFlag(said.hidden, "row: hidden"),
    collapsed: requireFlag(said.collapsed, "row: collapsed"),
  };
};

/**
 * A row's opening tag. The attributes sit in the order `CT_Row` declares them,
 * and each rides along only where it says something.
 * @param {number} at
 * @param {ReturnType<typeof requireRow>} row
 */
const rowOpen = (at, row) =>
  `<row r="${at}"` +
  (row.hidden ? ' hidden="1"' : "") +
  (row.level === 0 ? "" : ` outlineLevel="${row.level}"`) +
  (row.collapsed ? ' collapsed="1"' : "") +
  ">";

/**
 * The sheet's `sheetFormatPr`, or nothing for a sheet that outlines no row.
 * The format requires a default row height beside the level count; 15 is
 * Excel's own for the default 11-point font.
 * @param {number} deepest
 */
const formatXml = (deepest) =>
  deepest === 0 ? "" : `<sheetFormatPr defaultRowHeight="15" outlineLevelRow="${deepest}"/>`;

/**
 * A `Date` as the serial Excel stores. An invalid one has no serial to give.
 * @param {Date} date
 * @param {string} at
 */
const dateSerial = (date, at) => {
  if (Number.isNaN(date.getTime()))
    throw RangeError(`${at}: expected a valid Date, got an invalid one`);
  return serial(date);
};

/**
 * The shared-string index of a cell's text, plain or rich.
 * @param {ReturnType<typeof strings>} sst
 * @param {unknown} value
 * @param {string} at
 */
const textIndex = (sst, value, at) => {
  if (typeof value === "string")
    return sst.intern("s" + value, () => `<si>${textNode(value)}</si>`);
  if (Array.isArray(value)) return sst.intern(richKey(value, at), () => richXml(value, at));
  throw TypeError(`${at}: unsupported cell value ${JSON.stringify(value)}`);
};

/**
 * The type attribute and `<v>` of a present cell value.
 * @param {ReturnType<typeof strings>} sst
 * @param {unknown} value
 * @param {string} where
 * @returns {[string, string | number]}
 */
const cellBody = (sst, value, where) => {
  const at = `${where}.value`;
  if (typeof value === "number") return ["", number(requireNumber(value, at))];
  if (typeof value === "boolean") return [' t="b"', Number(value)];
  if (value instanceof Date) return ["", number(dateSerial(value, at))];
  return [' t="s"', textIndex(sst, value, at)];
};

/**
 * @param {ReturnType<typeof strings>} sst
 * @param {unknown} value
 * @param {number} styleId
 * @param {string} ref
 * @param {string} where
 */
const cellXml = (sst, value, styleId, ref, where) => {
  const s = styleId === 0 ? "" : ` s="${styleId}"`;
  if (absent(value)) return s === "" ? "" : `<c r="${ref}"${s}/>`;
  const [type, v] = cellBody(sst, value, where);
  return `<c r="${ref}"${s}${type}><v>${v}</v></c>`;
};

// A style that is not an object passes through untouched, so the interning
// below is still the one place that refuses it.
/** @param {unknown} declared */
const usable = (declared) => absent(declared) || typeof declared === "object";

// A declared style that names no number format of its own.
/** @param {Record<string, unknown> | null | undefined} declared */
const undated = (declared) => usable(declared) && absent(declared?.numberFormat);

/**
 * A date with no format of its own gets the short-date built-in, so a `Date`
 * reads back as a day rather than as the number underneath it.
 * @param {unknown} value
 * @param {Record<string, unknown> | null | undefined} declared
 */
const cellStyle = (value, declared) =>
  value instanceof Date && undated(declared) ? { ...declared, numberFormat: SHORT_DATE } : declared;

/** @param {number} frozen */
const paneXml = (frozen) =>
  frozen === 0
    ? '<sheetView workbookViewId="0"/>'
    : `<sheetView workbookViewId="0"><pane ySplit="${frozen}" topLeftCell="A${frozen + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${frozen + 1}" sqref="A${frozen + 1}"/></sheetView>`;

/**
 * @typedef {{ top: number, left: number, bottom: number, right: number }} Range
 */

/** @param {Range} range */
const rangeRef = (range) =>
  `${letters(range.left)}${range.top}:${letters(range.right)}${range.bottom}`;

/** @param {Range} a @param {Range} b */
const overlaps = (a, b) =>
  a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;

/**
 * A `<col>` per column that was given a width, in column order, and nothing
 * at all when none was. One element apiece rather than a `min`/`max` run: a
 * run may never span a column left unset, and the sparse case breaks that
 * first, for bytes on one of the smallest parts here.
 *
 * `customWidth` is not optional. Without it a reader treats the width as one
 * it derived, and is free to compute its own instead.
 *
 * @param {ReadonlyArray<number | null | undefined>} columns
 */
const colsXml = (columns) => {
  const body = columns
    .map((width, index) =>
      absent(width)
        ? ""
        : `<col min="${index + 1}" max="${index + 1}" width="${number(width)}" customWidth="1"/>`,
    )
    .join("");
  return body === "" ? "" : `<cols>${body}</cols>`;
};

// One rule of a filter's range, stated the way a caller reads it.
/** @param {boolean} ok @param {string} why */
const requireFilterRule = (ok, why) => {
  if (!ok) throw RangeError(`filter: ${why}`);
};

// A range may only cover cells that were written. `merge` and `filter` both
// ask it, of a row and of a column.
/** @param {number} at @param {number} written @param {string} what @param {string} where */
const requireWritten = (at, written, what, where) => {
  if (at > written)
    throw RangeError(`${where}: ${what} ${at} does not exist yet (the sheet has ${written})`);
};

/**
 * The rectangle an autofilter covers, 1-based and inclusive on all four
 * sides. A range whose bottom row was never written is a caller's bug, so
 * this takes the row count to check it against.
 *
 * @param {unknown} range
 * @param {number} rows how many rows the sheet holds
 * @param {number} columns how many columns it has written
 */
const requireFilter = (range, rows, columns) => {
  const rect = requireRecord(range, "filter", "{ top, left, bottom, right }");
  const top = requireIndex(rect.top, MAX_ROW, "filter: top");
  const left = requireIndex(rect.left, MAX_COLUMN, "filter: left");
  const bottom = requireIndex(rect.bottom, MAX_ROW, "filter: bottom");
  const right = requireIndex(rect.right, MAX_COLUMN, "filter: right");
  requireFilterRule(bottom >= top, `bottom ${bottom} is above top ${top}`);
  requireFilterRule(right >= left, `right ${right} is left of left ${left}`);
  requireWritten(bottom, rows, "row", "filter");
  requireWritten(right, columns, "column", "filter");
  return { top, left, bottom, right };
};

/**
 * @typedef {{ url: string } | { location: string }} LinkTarget
 * @typedef {{ ref: string } & LinkTarget} Link
 */

// A link names one of the two: a destination outside the workbook, or a
// reference inside it. Both at once names two destinations and neither names
// none, so the count is what this checks rather than each key in turn.
/** @param {unknown} target @returns {LinkTarget} */
const requireTarget = (target) => {
  const url = /** @type {any} */ (target)?.url;
  const location = /** @type {any} */ (target)?.location;
  if ((url === undefined) === (location === undefined))
    throw RangeError(
      `link: expected exactly one of url and location, got ${JSON.stringify(target)}`,
    );
  if (url === undefined) return { location: requireDestination(location, "link: location") };
  return { url: requireDestination(url, "link: url") };
};

// A destination is a string with something in it: an empty one is a link a
// reader opens onto nothing, which is a caller's mistake rather than a link.
/** @param {unknown} value @param {string} where */
const requireDestination = (value, where) => {
  const text = requireString(value, where);
  if (text === "") throw RangeError(`${where}: expected a destination, got an empty string`);
  return text;
};

// The destinations outside the workbook, deduplicated and in first-seen
// order: one relationship each, however many cells point at it. The same walk
// answers for the sheet part and for its relationships, so the two cannot
// disagree about which id a link carries.
/** @param {ReadonlyArray<Link>} links */
const externals = (links) => [
  ...new Set(links.filter((link) => "url" in link).map((link) => /** @type {any} */ (link).url)),
];

// Every link of one sheet, in call order. An external one points at a
// relationship of this worksheet's own part; an internal one carries its
// reference and needs none. `first` is the id the relationships start at,
// which is after the drawing where the sheet has one.
/** @param {ReadonlyArray<Link>} links @param {number} first */
const hyperlinksXml = (links, first) => {
  if (links.length === 0) return "";
  const targets = externals(links);
  const one = (/** @type {Link} */ link) =>
    "url" in link
      ? `<hyperlink ref="${link.ref}" r:id="rId${first + targets.indexOf(link.url)}"/>`
      : `<hyperlink ref="${link.ref}" location="${esc(link.location)}"/>`;
  return `<hyperlinks>${links.map(one).join("")}</hyperlinks>`;
};

/** @param {ReadonlyArray<Range>} merges */
const mergesXml = (merges) =>
  merges.length === 0
    ? ""
    : `<mergeCells count="${merges.length}">` +
      merges.map((range) => `<mergeCell ref="${rangeRef(range)}"/>`).join("") +
      `</mergeCells>`;

/**
 * The last column a merge reaches, once its width is known to be a count of
 * two or more that stays on the sheet.
 * @param {unknown} width
 * @param {number} at
 */
const mergeEnd = (width, at) => {
  if (!Number.isInteger(width) || /** @type {number} */ (width) < 2)
    throw RangeError(`merge: expected a width of 2 or more, got ${JSON.stringify(width)}`);
  const right = at + /** @type {number} */ (width) - 1;
  if (right > MAX_COLUMN) throw RangeError(`merge: the range ends past column ${MAX_COLUMN}`);
  return right;
};

/**
 * One entry of a width list. Absent is how a caller leaves a column alone, so
 * it is the one value that checks nothing.
 *
 * @param {unknown} width
 * @param {number} index 0-based, so the message can name the column
 */
const checkWidth = (width, index) => {
  if (absent(width)) return;
  const chars = requireNumber(width, `widths: column ${index + 1}`);
  if (chars > 0 && chars <= MAX_WIDTH) return;
  throw RangeError(
    `widths: column ${index + 1} expected a width above 0 and at most ${MAX_WIDTH}, got ${chars}` +
      ` -- null leaves a column unset`,
  );
};

/** @param {unknown} margin */
const requireMargin = (margin) => {
  const points = requireNumber(margin, "print: margin");
  if (points < 0) throw RangeError(`print: margin cannot be negative, got ${points}`);
};

/** @param {unknown} size */
const requirePaper = (size) => {
  if (!PAPER.has(/** @type {string} */ (size)))
    throw RangeError(
      `print: unknown paper size ${JSON.stringify(size)} -- known: ${[...PAPER.keys()].join(", ")}`,
    );
};

// Zero clears, as it does for a freeze, so a caller that set a count has a way
// back. Not `requireIndex`: its message names 1 as the floor, and here it is 0.
/** @param {unknown} rows */
const requireTitles = (rows) => {
  if (rows === 0) return;
  if (!isIndex(rows, MAX_ROW))
    throw RangeError(
      `print: titles expected 0, or a row count between 1 and ${MAX_ROW}, got ${JSON.stringify(rows)}`,
    );
};

/** @param {string} key @returns {(text: unknown) => void} */
const requireFurniture = (key) => (text) => {
  const stored = furniture(text, `print: ${key}`);
  if (stored.length > MAX_FURNITURE)
    throw RangeError(
      `print: ${key} stores as ${stored.length} characters, and a reader holds at most ${MAX_FURNITURE}`,
    );
};

/** @param {unknown} orientation */
const requireOrientation = (orientation) => {
  if (!ORIENTATION.has(/** @type {string} */ (orientation)))
    throw RangeError(`print: expected portrait or landscape, got ${JSON.stringify(orientation)}`);
};

/** @param {number} width @param {number} height */
const requireSize = (width, height) => {
  if (width <= 0 || height <= 0)
    throw RangeError(`place: expected a positive size, got ${width} by ${height}`);
};

// Every print key this surface checks, and the check it gets. Key order is
// the order a caller hears about a mistake. A table rather than a run of
// guards: a guard apiece puts `print` over the cyclomatic ceiling. `fit` is
// absent because only `true` writes anything.
const PRINT_CHECKS = {
  margin: requireMargin,
  size: requirePaper,
  orientation: requireOrientation,
  titles: requireTitles,
  header: requireFurniture("header"),
  footer: requireFurniture("footer"),
  firstHeader: requireFurniture("firstHeader"),
  firstFooter: requireFurniture("firstFooter"),
};

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
  /** @type {Link[]} */
  const links = [];
  let frozen = 0;
  let widest = 1;
  // The deepest outline level any row carries; 0 writes no `sheetFormatPr`.
  let deepest = 0;
  // This sheet's `autoFilter` element, or the empty string for a sheet that
  // asked for none. Rendered where it is set, the way a row is.
  let filtered = "";
  // Widths by position, the first entry being column A. Replaced wholesale by
  // `widths`, never merged: a hole in a positional list cannot mean both
  // "leave this one alone" and "clear it".
  /** @type {ReadonlyArray<number | null | undefined>} */
  let columns = [];
  // What `print` was told, or null. Nothing reaches the file until a caller
  // asks: a reader's own print defaults are better than this writer guessing.
  /** @type {PrintSetup | null} */
  let printing = null;

  /**
   * One cell of a row: the empty string for a gap, else its `<c>`.
   * @param {unknown} cell
   * @param {number} index 0-based column
   * @param {number} at 1-based row
   */
  const cellAt = (cell, index, at) => {
    if (absent(cell)) return "";
    if (typeof cell !== "object")
      throw TypeError(`row: cell ${index + 1} is not a { value, style } object`);
    const where = `row ${at}, cell ${index + 1}`;
    const { value, style } = /** @type {Record<string, unknown>} */ (cell);
    const declared = /** @type {Record<string, unknown> | null | undefined} */ (style);
    const styleId = styles.intern(cellStyle(value, declared), `${where}.style`);
    return cellXml(sst, value, styleId, letters(index + 1) + at, where);
  };

  return {
    name,

    /**
     * @param {ReadonlyArray<unknown>} cells
     * @param {unknown} [options]
     */
    row(cells, options) {
      if (!Array.isArray(cells))
        throw TypeError(`row: expected an array of cells, got ${JSON.stringify(cells)}`);
      const at = rows.length + 1;
      if (cells.length > MAX_COLUMN)
        throw RangeError(`row: a sheet holds at most ${MAX_COLUMN} columns`);
      const row = requireRow(options);
      deepest = Math.max(deepest, row.level);

      let body = "";
      for (let index = 0; index < cells.length; index++) body += cellAt(cells[index], index, at);
      widest = Math.max(widest, cells.length);
      rows.push(rowOpen(at, row) + body + "</row>");
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
      requireWritten(row, rows.length, "row", "merge");
      const range = { top: row, left: at, bottom: row, right: mergeEnd(width, at) };
      const clash = merges.find((other) => overlaps(range, other));
      if (clash !== undefined)
        throw RangeError(`merge: the range ${rangeRef(range)} overlaps ${rangeRef(clash)}`);
      merges.push(range);
      widest = Math.max(widest, range.right);
    },

    /**
     * Link one cell. `target` names exactly one of `url`, a destination
     * outside this workbook, and `location`, a reference inside it -- a cell
     * as `'Sheet2'!A1`, or a defined name.
     *
     * Linking writes no cell: what a reader shows is whatever the row already
     * put there, which is why the row must exist first, as a merge's must. A
     * reader holds one link per cell, so a second on the same cell throws
     * rather than quietly replacing the first.
     *
     * @param {number} row 1-based
     * @param {number} at 1-based column
     * @param {LinkTarget} target
     */
    link(row, at, target) {
      requireIndex(row, MAX_ROW, "link: row");
      requireIndex(at, MAX_COLUMN, "link: at");
      requireWritten(row, rows.length, "row", "link");
      const ref = `${letters(at)}${row}`;
      if (links.some((other) => other.ref === ref))
        throw RangeError(`link: ${ref} already carries a link`);
      links.push({ ref, ...requireTarget(target) });
    },

    get links() {
      return links;
    },

    /**
     * How this worksheet prints. Every key is optional, and a worksheet that
     * never calls this carries no print setup at all: a reader's own defaults
     * beat a guess.
     *
     * @param {PrintSetup} setup
     */
    print(setup) {
      const declared = requireRecord(setup, "print", "a setup object");
      for (const [key, check] of Object.entries(PRINT_CHECKS))
        if (declared[key] !== undefined) check(declared[key]);
      printing = { ...printing, ...setup };
    },

    // Read by the workbook part rather than by this sheet's own: a print title
    // is a defined name, and defined names live in `xl/workbook.xml`.
    get titles() {
      return printing?.titles ?? 0;
    },

    /** @param {number} count */
    freeze(count) {
      if (!Number.isInteger(count) || count < 0 || count >= MAX_ROW)
        throw RangeError(`freeze: expected a row count of 0 or more, got ${JSON.stringify(count)}`);
      frozen = count;
    },

    /**
     * Column widths by position, the first entry being column A. `null` leaves
     * a column unset, so a reader keeps its own default for it.
     *
     * Replaces rather than merges, and an empty list clears. A width never
     * widens the sheet: `dimension` describes the cells that were written, and
     * sizing a column writes no cell.
     *
     * @param {ReadonlyArray<unknown>} list
     */
    widths(list) {
      if (!Array.isArray(list))
        throw TypeError(`widths: expected an array of widths, got ${JSON.stringify(list)}`);
      if (list.length > MAX_COLUMN)
        throw RangeError(`widths: a sheet holds at most ${MAX_COLUMN} columns`);
      list.forEach(checkWidth);
      columns = list.slice();
    },

    /**
     * Put an autofilter over a range, 1-based and inclusive. A worksheet takes
     * one, so this replaces rather than merges, and `null` clears it.
     *
     * @param {unknown} range
     */
    filter(range) {
      // Excel also records an autofilter as a sheet-scoped
      // `_xlnm._FilterDatabase` defined name; this writes none. Measured in
      // both readers: Excel opens such a file with no repair prompt and
      // reports the filter as on, and LibreOffice writes the name itself on
      // save. The name is a reader's bookkeeping.
      filtered =
        range === null
          ? ""
          : `<autoFilter ref="${rangeRef(requireFilter(range, rows.length, widest))}"/>`;
    },

    /**
     * @param {unknown} id
     * @param {{ row: number, col?: number, width: number, height: number }} at
     */
    place(id, at) {
      if (!knows(id)) throw TypeError(`place: no image with id ${JSON.stringify(id)}`);
      requireRecord(at, "place", "{ row, col, width, height }");
      const row = requireIndex(at.row, MAX_ROW, "place: row");
      const col = requireIndex(at.col === undefined ? 1 : at.col, MAX_COLUMN, "place: col");
      const width = requireNumber(at.width, "place: width");
      const height = requireNumber(at.height, "place: height");
      requireSize(width, height);
      pictures.push({ id: /** @type {number} */ (id), row, col, width, height });
    },

    get pictures() {
      return pictures;
    },

    /** @param {number | null} drawing the relationship id of this sheet's drawing part */
    xml(drawing) {
      const dimension = `A1:${letters(widest)}${Math.max(rows.length, 1)}`;

      // `fitToPage` lives on `sheetPr`, which the schema puts before every
      // other child of a worksheet.
      const properties =
        printing?.fit === true ? `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` : "";

      return (
        DECLARATION +
        `<worksheet xmlns="${NS}" xmlns:r="${NS_R}">` +
        properties +
        `<dimension ref="${dimension}"/>` +
        `<sheetViews>${paneXml(frozen)}</sheetViews>` +
        formatXml(deepest) +
        colsXml(columns) +
        `<sheetData>${rows.join("")}</sheetData>` +
        filtered +
        mergesXml(merges) +
        hyperlinksXml(links, drawing === null ? 1 : drawing + 1) +
        printXml(printing) +
        (drawing === null ? "" : `<drawing r:id="rId${drawing}"/>`) +
        `</worksheet>`
      );
    },
  };
};

/**
 * The drawing part for one sheet: a `oneCellAnchor` per placement, anchored
 * to the top-left of its cell. The XML counts rows and columns from zero and
 * the subtraction happens here, so the surface stays 1-based throughout.
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

/** @param {Uint8Array} a @param {Uint8Array} b of the same length */
const sameBytes = (a, b) => {
  for (let at = 0; at < a.length; at++) if (a[at] !== b[at]) return false;
  return true;
};

// The same array twice is the common case, and the one the byte loop is worst
// at: equal bytes never exit early, so it reads the whole image every time. A
// caller placing one logo on a thousand rows pays 54 ms for that, and nothing
// for the identity check.
/** @param {Uint8Array} a @param {Uint8Array} b */
const same = (a, b) => a === b || (a.length === b.length && sameBytes(a, b));

/**
 * @param {unknown} name
 * @returns {string}
 */
const requireSheetName = (name) => {
  if (typeof name !== "string" || name === "")
    throw RangeError(`sheet: expected a name, got ${JSON.stringify(name)}`);
  if (name.length > NAME_LIMIT)
    throw RangeError(`sheet: a name is at most ${NAME_LIMIT} characters, got ${name.length}`);
  return name;
};

/**
 * @param {number} sheets how many worksheets the workbook holds
 * @param {number} drawings how many of them carry a drawing part
 * @param {ReadonlyArray<{ format: string }>} media
 */
const contentTypesXml = (sheets, drawings, media) => {
  const overrides = Array.from(
    { length: sheets },
    (_sheet, index) =>
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="${CT}.spreadsheetml.worksheet+xml"/>`,
  ).join("");
  const drawingOverrides = Array.from(
    { length: drawings },
    (_each, at) =>
      `<Override PartName="/xl/drawings/drawing${at + 1}.xml" ContentType="${CT}.drawing+xml"/>`,
  ).join("");
  const defaults = [...new Set(media.map((each) => each.format))]
    .map((format) => `<Default Extension="${format}" ContentType="${FORMATS.get(format)}"/>`)
    .join("");

  return (
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
    `</Types>`
  );
};

const ROOT_RELS =
  DECLARATION +
  `<Relationships xmlns="${NS_PKG_REL}">` +
  `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
  `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
  `</Relationships>`;

const APP_XML =
  DECLARATION +
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>werkmap</Application></Properties>`;

/** @param {string} name @param {string | undefined} value */
const tag = (name, value) => (value === undefined ? "" : `<${name}>${esc(value)}</${name}>`);

/**
 * @param {{ title?: string, creator?: string, subject?: string, description?: string }} meta
 */
const coreXml = (meta) =>
  DECLARATION +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
  tag("dc:title", meta.title) +
  tag("dc:subject", meta.subject) +
  tag("dc:creator", meta.creator) +
  tag("dc:description", meta.description) +
  `<dcterms:created xsi:type="dcterms:W3CDTF">${CREATED}</dcterms:created>` +
  `<dcterms:modified xsi:type="dcterms:W3CDTF">${CREATED}</dcterms:modified>` +
  `</cp:coreProperties>`;

// A sheet reference inside a defined name is always quoted, so a name holding
// a space or a symbol needs no special case. An apostrophe in the name closes
// the quote early unless it doubles, and the whole reference is XML-escaped
// after that like every other name this writer emits.
/** @param {string} name @param {number} rows */
const titlesRef = (name, rows) => `${esc(`'${name.replace(/'/g, "''")}'`)}!$1:$${rows}`;

/**
 * The built-in name that repeats the top rows of a sheet on every printed
 * page. `localSheetId`, the sheet's 0-based position in `<sheets>`, is what
 * scopes the name to that worksheet rather than to the whole workbook.
 *
 * @param {ReadonlyArray<{ name: string, titles: number }>} sheets
 */
const definedNamesXml = (sheets) => {
  const names = sheets
    .map((sheet, index) =>
      sheet.titles === 0
        ? ""
        : `<definedName name="_xlnm.Print_Titles" localSheetId="${index}">${titlesRef(sheet.name, sheet.titles)}</definedName>`,
    )
    .join("");
  return names === "" ? "" : `<definedNames>${names}</definedNames>`;
};

/** @param {ReadonlyArray<{ name: string, titles: number }>} sheets */
const workbookXml = (sheets) =>
  DECLARATION +
  `<workbook xmlns="${NS}" xmlns:r="${NS_R}"><sheets>` +
  sheets
    .map(
      (sheet, index) =>
        `<sheet name="${esc(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
    )
    .join("") +
  `</sheets>` +
  definedNamesXml(sheets) +
  `</workbook>`;

/** @param {number} sheets how many worksheets the workbook holds */
const workbookRelsXml = (sheets) =>
  DECLARATION +
  `<Relationships xmlns="${NS_PKG_REL}">` +
  Array.from(
    { length: sheets },
    (_sheet, index) =>
      `<Relationship Id="rId${index + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
  ).join("") +
  `<Relationship Id="rId${sheets + 1}" Type="${REL}/styles" Target="styles.xml"/>` +
  `<Relationship Id="rId${sheets + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>` +
  `</Relationships>`;

/**
 * @typedef {(name: string, body: string | Uint8Array, store?: boolean) => void} Part
 * @typedef {{ sheet: ReturnType<typeof worksheet>, index: number }} Drawing
 */

// What one worksheet's own part points at: its drawing, where it has one, and
// then a relationship per destination outside the workbook. The drawing leads
// because it did before links existed, and `hyperlinksXml` counts from the
// same place.
/** @param {number} at @param {ReadonlyArray<Link>} links */
const sheetRels = (at, links) =>
  DECLARATION +
  `<Relationships xmlns="${NS_PKG_REL}">` +
  (at === -1
    ? ""
    : `<Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing${at + 1}.xml"/>`) +
  externals(links)
    .map(
      (url, index) =>
        `<Relationship Id="rId${(at === -1 ? 1 : 2) + index}" Type="${REL}/hyperlink" Target="${esc(url)}" TargetMode="External"/>`,
    )
    .join("") +
  `</Relationships>`;

/**
 * Every worksheet part, and the relationship part of each sheet that points
 * at anything of its own — a drawing, a link, or both.
 * @param {Part} part
 * @param {ReadonlyArray<ReturnType<typeof worksheet>>} sheets
 * @param {ReadonlyArray<Drawing>} drawings
 */
const sheetParts = (part, sheets, drawings) => {
  for (const [index, sheet] of sheets.entries()) {
    const at = drawings.findIndex((each) => each.index === index);
    part(`xl/worksheets/sheet${index + 1}.xml`, sheet.xml(at === -1 ? null : 1));
    if (at !== -1 || externals(sheet.links).length > 0)
      part(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, sheetRels(at, sheet.links));
  }
};

/**
 * Each drawing part and its relationships. A drawing numbers its own
 * relationships from one, in the order its pictures were placed, so they line
 * up with its own part rather than with the workbook-wide media list.
 * @param {Part} part
 * @param {ReadonlyArray<Drawing>} drawings
 * @param {ReadonlyArray<{ format: string }>} media
 */
const drawingParts = (part, drawings, media) => {
  for (const [at, each] of drawings.entries()) {
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
};

/**
 * Open a workbook.
 *
 * @param {{ title?: string, creator?: string, subject?: string, description?: string }} [meta]
 */
export const workbook = (meta = {}) => {
  requireRecord(meta, "workbook", "a metadata object");

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

  const properties = {
    title: property("title"),
    creator: property("creator"),
    subject: property("subject"),
    description: property("description"),
  };

  /** @param {unknown} id */
  const knows = (id) =>
    Number.isInteger(id) &&
    /** @type {number} */ (id) >= 0 &&
    /** @type {number} */ (id) < media.length;

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
      const seen = media.findIndex((each) => each.format === format && same(each.bytes, bytes));
      if (seen !== -1) return seen;
      media.push({ bytes, format });
      return media.length - 1;
    },

    /** @param {string} name */
    sheet(name) {
      requireSheetName(name);
      if (FORBIDDEN_IN_NAME.test(name))
        throw RangeError(
          `sheet: a name cannot contain : \\ / ? * [ ], got ${JSON.stringify(name)}`,
        );
      if (sheets.some((each) => each.name === name))
        throw RangeError(`sheet: a sheet named ${JSON.stringify(name)} already exists`);
      const made = worksheet(name, styles, sst, knows);
      sheets.push(made);
      return made;
    },

    async bytes() {
      if (sheets.length === 0) throw RangeError("bytes: a workbook needs at least one sheet");

      /** @type {Array<{ name: string, body: Uint8Array, store?: boolean }>} */
      const entries = [];
      /** @type {Part} */
      const part = (name, body, store) =>
        entries.push({
          name,
          body: store ? /** @type {Uint8Array} */ (body) : utf8(/** @type {string} */ (body)),
          store,
        });

      // The sheets that carry a drawing part, in sheet order: drawing parts
      // are numbered by this list, not by sheet index.
      const drawings = sheets
        .map((sheet, index) => ({ sheet, index }))
        .filter((each) => each.sheet.pictures.length > 0);

      part("[Content_Types].xml", contentTypesXml(sheets.length, drawings.length, media));
      part("_rels/.rels", ROOT_RELS);
      part("docProps/core.xml", coreXml(properties));
      part("docProps/app.xml", APP_XML);
      part("xl/workbook.xml", workbookXml(sheets));
      part("xl/_rels/workbook.xml.rels", workbookRelsXml(sheets.length));
      sheetParts(part, sheets, drawings);
      part("xl/styles.xml", styles.xml());
      part("xl/sharedStrings.xml", sst.xml());
      drawingParts(part, drawings, media);
      for (const [index, each] of media.entries())
        part(`xl/media/image${index + 1}.${each.format}`, each.bytes, true);

      return zip(entries);
    },
  };
};
