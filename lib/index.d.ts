/** A colour as `#rrggbb`, or the `#rgb` shorthand. There is no alpha channel. */
export type Colour = string;

/** A font, written in full wherever it appears: an OOXML run inherits nothing. */
export interface Font {
  /** Defaults to `Calibri`. */
  name?: string;
  /** In points. Defaults to 11. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: Colour;
}

/** One side of a cell border. */
export interface Edge {
  style: "thin" | "dashed" | "dotted";
  color?: Colour;
}

export interface Border {
  left?: Edge | null;
  right?: Edge | null;
  top?: Edge | null;
  bottom?: Edge | null;
}

export interface Alignment {
  horizontal?: "left" | "center" | "right" | "justify" | "fill";
  vertical?: "top" | "center" | "bottom" | "justify";
  wrapText?: boolean;
}

export interface Style {
  font?: Font | null;
  /** A fill *is* a colour: solid is the only pattern this writer emits. */
  fill?: Colour | null;
  border?: Border | null;
  alignment?: Alignment | null;
  /**
   * A number format code, spelled out. A code matching a built-in is written
   * as that built-in's id; anything else is interned.
   */
  numberFormat?: string | null;
}

/** One stretch of text inside a rich-text cell. */
export interface Run {
  text: string;
  font?: Font | null;
}

/**
 * A cell value. `null` is an empty cell, which is still written when the cell
 * carries a style. Rich text is a bare array of runs.
 */
export type Value = string | number | boolean | Date | readonly Run[] | null;

export interface Cell {
  value?: Value;
  style?: Style | null;
}

/** Where a picture sits, and how large it is drawn. */
export interface Placement {
  /** 1-based row of the cell the picture anchors to. */
  row: number;
  /** 1-based column of that cell. Defaults to 1. */
  col?: number;
  /** CSS pixels at 96 dpi. */
  width: number;
  /** CSS pixels at 96 dpi. */
  height: number;
}

/** How a worksheet prints. Every key is optional. */
/**
 * Text for a page header or footer: a string, or parts in order, where a
 * `{ field }` is a value the reader fills in as it paginates — `page` is the
 * page number and `pages` the page count.
 */
/**
 * One part of a page header or footer: text, a `{ field }` the reader fills
 * in as it paginates (`page` is the page number, `pages` the page count), or
 * `{ text }` in a look — bold, and a font size from 1 to 99 — that holds
 * until a later part changes it.
 */
export type PrintPart =
  | string
  | { field: "page" | "pages" }
  | { text: string; bold?: boolean; size?: number };

/** One section of a header: a string, or parts in order. */
export type PrintSection = string | readonly PrintPart[];

/**
 * Text for a page header or footer: one section, printed on the left, or the
 * three sections the format has, each optional.
 */
export type PrintText =
  | PrintSection
  | { left?: PrintSection; center?: PrintSection; right?: PrintSection };

export interface PrintSetup {
  /** All four page margins, in points. Defaults to the reader's own. */
  margin?: number;
  /** A paper size OOXML names. There is no arbitrary width and height. */
  size?: "letter" | "tabloid" | "legal" | "A3" | "A4" | "A5";
  orientation?: "portrait" | "landscape";
  /** Scale the sheet to one page wide, and as many pages tall as it takes. */
  fit?: boolean;
  /**
   * Repeat the top `rows` rows at the top of every printed page. `0` clears,
   * the way a `freeze` of `0` does. Scoped to this worksheet alone.
   */
  titles?: number;
  /**
   * Text printed at the top of every page: one section on the left, or the
   * three sections named. Text is text: an ampersand is an ampersand and a
   * newline a line break. The format's own codes are not on this surface;
   * the page number, the page count, bold and a size are named as parts
   * instead. At most 255 characters as stored.
   */
  header?: PrintText;
  /** Text printed at the bottom of every page. Same rules as `header`. */
  footer?: PrintText;
  /**
   * The first page's header, where it differs from the rest. Naming either
   * first-page part makes the first page different; a part not named prints
   * nothing there.
   */
  firstHeader?: PrintText;
  /** The first page's footer, where it differs from the rest. */
  firstFooter?: PrintText;
}

/** The rectangle an autofilter covers. 1-based, and inclusive on all four sides. */
export interface FilterRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** What a row says about itself beyond its cells. */
export interface RowOptions {
  /**
   * The row's outline level, an integer from 0 to 7. A reader draws the
   * levels as collapsible groups in its margin, with the summary row read as
   * the one below the group. `0`, the default, is a row outside any group.
   */
  level?: number;
  /**
   * Whether the row is hidden. A collapsed outline group is its content rows
   * hidden and its summary row `collapsed`; hiding alone leaves the group's
   * control showing expanded over rows nobody can see.
   */
  hidden?: boolean;
  /**
   * Whether this row carries the control of a collapsed group. It goes on the
   * **summary** row — the one below the group — not on the hidden rows.
   */
  collapsed?: boolean;
}

export interface Sheet {
  /** This sheet's name. */
  readonly name: string;
  /**
   * Append a row. `null` is an empty unstyled cell. Returns the row's 1-based
   * position, which `merge` and `place` take. `options.level` puts the row at
   * an outline level, and `hidden` / `collapsed` say whether it is shown and
   * whether it carries a collapsed group's control. An option this does not
   * know is refused rather than ignored.
   */
  row(cells: readonly (Cell | null)[], options?: RowOptions): number;
  /**
   * Merge `width` columns of `row`, starting at the 1-based column `at`.
   * Throws on a width below 2, an overlap, or a row that does not exist yet.
   */
  merge(row: number, at: number, width: number): void;
  /** Freeze the top `count` rows. `0` clears. */
  freeze(count: number): void;
  /**
   * Column widths by position: the first entry is column A, and `null` leaves
   * a column unset so a reader keeps its own default for it.
   *
   * The unit is the format's own — a count of characters of the workbook's
   * default font, the number Excel's column-width box shows. A width is above
   * 0 and at most 255.
   *
   * Replaces rather than merges, and an empty list clears.
   */
  widths(list: readonly (number | null)[]): void;
  /**
   * How this worksheet prints. A worksheet that never calls this carries no
   * print setup at all, so a reader applies its own defaults. Calls merge, so
   * two calls naming different keys both take effect.
   */
  print(setup: PrintSetup): void;
  /**
   * Put an autofilter over a range, 1-based and inclusive on all four sides.
   * A worksheet takes one, so this replaces rather than merges, and `null`
   * clears it. The range's bottom row and rightmost column must already have
   * been written, checked when you call.
   */
  filter(range: FilterRange | null): void;
  /** Float a picture over the sheet, anchored to one cell. */
  place(id: number, at: Placement): void;
}

export interface Metadata {
  title?: string;
  creator?: string;
  subject?: string;
  description?: string;
}

export interface Workbook {
  /**
   * Embed an image. Identical bytes deduplicate and return the id already
   * issued. The id is what `Sheet.place` takes.
   */
  image(bytes: Uint8Array, format: "png" | "jpeg"): number;
  /**
   * Add a worksheet. Throws on a name Excel refuses: empty, over 31
   * characters, containing `: \ / ? * [ ]`, or a duplicate.
   */
  sheet(name: string): Sheet;
  /**
   * The finished package. Not terminal in the sense of sealing the document:
   * call it as often as you like, and the same call sequence yields the same
   * bytes.
   */
  bytes(): Promise<Uint8Array>;
}

/**
 * Open a workbook.
 *
 * The document's created and modified dates are always `1970-01-01T00:00:00Z`
 * and cannot be overridden: byte-identity for the same input is this writer's
 * headline promise, and an overridable clock would redefine "the same input"
 * as "the same input at the same wall time".
 */
export declare const workbook: (meta?: Metadata) => Workbook;
