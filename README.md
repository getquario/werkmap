# werkmap

A tiny spreadsheet writer for JavaScript. It writes an `.xlsx` file — rows, typed cells, styles, merged ranges, a frozen header, outline levels, floating images — and nothing else. It does not read one. _Werkmap_ is Dutch for a workbook, which is the one thing this package makes.

Writing is the whole surface, so the package stays small enough to audit. The container is written against the ZIP specification over `CompressionStream`, every part is a string, and no value ever turns into code — so it runs unchanged under a Content Security Policy with no `unsafe-*` of any kind, which is what a spreadsheet export in the browser usually cannot do.

- **Byte-identical output.** Nothing consults a clock, a locale or a random source, and the ZIP epoch is pinned, so two renders of the same report are the same bytes and a build that caches by content hash keeps working.
- **Foreign readers accept it.** The suite loads every workbook it writes back through ExcelJS, an independent implementation — so the tests prove a real reader opens the file, not just that the bytes look plausible.
- **Zero dependencies.** 7.7 kB minified and brotlied, for the whole writer.
- **Strict CSP, including in the browser.** No `unsafe-eval` and no `unsafe-inline`. A Chromium page under `default-src 'none'; script-src 'self'` writes a workbook and reports any violation back, and the Node suite runs on `--disallow-code-generation-from-strings`.
- **Write-only, deliberately.** No reader, no formula engine, no chart support — see [Is werkmap the right tool?](#is-werkmap-the-right-tool) before you install it.
- **Hardened.** 85 tests at 100% branch coverage.

```js
import { workbook } from "werkmap";

const wb = workbook({ title: "Sales" });
const sheet = wb.sheet("Report");

sheet.row([{ value: "Product" }, { value: "Amount" }]);
sheet.row([{ value: "Laptop" }, { value: 1000, style: { numberFormat: "#,##0.00" } }]);

const bytes = await wb.bytes();
//=> Uint8Array — and running this again produces byte-identical output
```

<img src="https://getquario.com/favicon.svg" alt="Quario logo" width="16" height="16" /> <b>werkmap</b> is built by the team behind <b><a href="https://getquario.com?utm_source=github&utm_medium=readme&utm_campaign=werkmap">Quario</a></b>, a declarative reporting engine for JavaScript that renders JSON report definitions to <b>HTML, PDF, workbooks, and Word</b> — without <code>eval</code>.

## Contents

- [Install](#install)
- [Usage](#usage)
- [Is werkmap the right tool?](#is-werkmap-the-right-tool)
- [API](#api)
- [Styles](#styles)
- [Guarantees, not options](#guarantees-not-options)
- [Determinism](#determinism)
- [Errors](#errors)
- [Content Security Policy](#content-security-policy)
- [Environments](#environments)
- [Contributing](#contributing)
- [License](#license)

## Install

```bash
npm install werkmap
```

Node.js 22 or newer, ESM only. TypeScript declarations ship with the package; nothing extra to install.

## Usage

```js
import { workbook } from "werkmap";

const wb = workbook({ title: "Sales", creator: "Acme BV" });
const logo = wb.image(png, "png");
const sheet = wb.sheet("Report");

const head = sheet.row([
  { value: "Sales", style: { font: { bold: true, size: 16 } } },
  { value: null },
]);
sheet.merge(head, 1, 2);

sheet.row([{ value: "Product" }, { value: "Amount" }]);
sheet.row([{ value: "Laptop" }, { value: 1000, style: { numberFormat: "#,##0.00" } }]);
sheet.freeze(head);
sheet.place(logo, { row: 1, width: 120, height: 40 });

const bytes = await wb.bytes(); // Uint8Array
```

Writing the file belongs to you: `bytes()` hands back a `Uint8Array`, which Node writes with `fs` and a browser hands to a download.

## Is werkmap the right tool?

werkmap writes the slice of the format a report needs, and refuses the rest.

**It fits when:**

- You produce spreadsheets and never consume them.
- The output must be reproducible — a content hash, a golden test, a cache key.
- The code runs under a strict Content Security Policy, in a browser or a worker.
- You care what the dependency costs: there are none, and the whole writer is one readable file.

**It does not fit when:**

- You need to read, edit or convert an existing workbook.
- You need formulas, charts, pivot tables, conditional formatting, data validation, hyperlinks, comments, sheet protection, or a page header and footer.
- You need row heights. A reader sizes rows from its own defaults. Column widths it does write — see `sheet.widths`.
- You have more rows than fit in memory. There is no streaming, row-at-a-time output.

## API

### `workbook(meta?) -> Workbook`

`meta` is one optional bag of optional strings: `{ title, creator, subject, description }`.

Members: `image`, `sheet`, `bytes`.

### `workbook.image(bytes, format) -> number`

`bytes` is a `Uint8Array` and `format` is `"png"` or `"jpeg"`. Returns an id for `sheet.place`.

Identical bytes deduplicate and return the id already issued, so the same logo on twelve sheets is stored once.

### `workbook.sheet(name) -> Sheet`

Any number of sheets, in call order. Throws on a name a reader refuses: empty, over 31 characters, containing `: \ / ? * [ ]`, or a duplicate.

### `workbook.bytes() -> Promise<Uint8Array>`

The finished package. This does not seal the document: call it as often as you like, and the same call sequence yields the same bytes.

### `sheet.row(cells, options?) -> number`

`cells` is an array; the returned number is the row's 1-based position, which `merge`, `freeze` and `place` take.

`options.level` is the row's **outline level**, an integer from 0 to 7. A reader draws the levels as collapsible groups in its left margin — the way Excel's own Group command does — and reads the summary row as the one **below** each group, which is where a total row sits. `0`, the default, is a row outside any group, and a row given no options is at 0.

`options.hidden` and `options.collapsed` say how a group opens. A collapsed group is **both**: its content rows `hidden`, and its summary row `collapsed`. Hiding alone leaves the group's control showing expanded over rows nobody can see, which is a document in two minds rather than a collapsed group. Neither is set by default, so a row you write is in the document and open unless you say otherwise, and what the reader then does with the controls is theirs.

```js
sheet.row([{ value: "North" }], { level: 1 });
sheet.row([{ value: "Laptop" }], { level: 2, hidden: true });
sheet.row([{ value: "Subtotal" }], { level: 1, collapsed: true });
```

An option this does not know is **refused**, not ignored: a misspelt one would otherwise be a row attribute you asked for and never got.

A sheet that outlines any row also states the workbook's default row height, 15 points, because the format requires it beside the deepest level. A sheet with no outline states neither, and a reader keeps its own default. That is the one place this writer names a row height, and it names the default rather than one of its own.

Each element is `{ value, style }`, or `null` for an empty unstyled cell. `{ value: null, style }` is an empty **styled** cell, which is what a merged span's remaining columns need.

`value` is one of:

| Kind      | Shape                 | Written as                              |
| --------- | --------------------- | --------------------------------------- |
| text      | a `string`            | interned into the shared string table   |
| number    | a finite `number`     | a numeric cell                          |
| boolean   | `true` / `false`      | a boolean cell                          |
| date      | a `Date`              | a 1900-system serial, converted in UTC  |
| rich text | `[{ text, font }, …]` | one formatted run per entry             |
| empty     | `null`                | an empty cell, written only when styled |

Rich text is a **bare array** — there is no wrapper object, because nothing else in the value domain is an array. Each run carries its font in full, since a run inherits nothing from the cell font.

### `sheet.merge(row, at, width)`

One merged range across `width` columns of `row`, starting at the 1-based column `at`.

### `sheet.freeze(rows)`

Freeze the top `rows` rows. `0` clears.

### `sheet.widths(list)`

Column widths by position: the first entry is column A, and `null` leaves a column unset so a reader keeps its own default for it.

The unit is the format's own — **a count of characters of the workbook's default font**, which is the number Excel's column-width box shows. A width is above 0 and at most 255. Pixels would read more naturally beside `place`, but converting them runs through the default font's maximum digit width, and any cell here may name a font of its own, so that constant would be wrong for most workbooks. The awkward unit invents no number.

The list **replaces** rather than merging, and an empty list clears — a hole in a positional list cannot mean both "leave this one alone" and "clear it". Call it before or after the rows; a width never widens the sheet, because `dimension` describes the cells that were written and sizing a column writes no cell. A worksheet that never calls this carries **no `cols` element at all**.

What is written is the number you gave, verbatim. Excel may report a slightly different one after a round trip, because it re-derives a width from the default font's digit width — that is the reader's arithmetic, not this writer's.

There is no `hidden`, no column outline level and no per-column style: a zero width is the back door to a hidden column, so `0` throws and points at `null`.

### `sheet.filter(range)`

Put an autofilter over `{ top, left, bottom, right }` — 1-based, and inclusive on all four sides. It adds the dropdown controls and **hides nothing**: every row you wrote is still in the document, and what the reader does with the control is theirs.

A worksheet takes one, so a second call **replaces** the first, and `null` clears it. The range must already have been written — its bottom row and its rightmost column both, the way a merge's row must — because a filter over cells nobody wrote is a caller's mistake rather than an empty range. Both are checked when you call, so write the rows first. A worksheet that never calls this carries **no `autoFilter` element at all**.

Excel also records an autofilter as a sheet-scoped `_xlnm._FilterDatabase` defined name, and this writes none. That is measured rather than assumed, in both readers: Excel opens a file carrying only the element, with no repair prompt, and reports the sheet's autofilter as on; LibreOffice opens it, keeps the filter, and writes that name itself on save. The name is a reader's bookkeeping rather than something a file owes.

### `sheet.print(setup)`

How this worksheet prints: `{ margin, size, orientation, fit, titles, header, footer }`, every key optional. `margin` is all four page margins in points, `size` one of `letter`, `tabloid`, `legal`, `A3`, `A4`, `A5`, `orientation` either `portrait` or `landscape`, and `fit: true` scales the sheet to one page wide and as many pages tall as it takes.

A worksheet that never calls this carries **no print setup at all**, so a reader applies its own defaults rather than this writer's opinion. Calls merge, so two calls naming different keys both take effect. Print setup is per worksheet, which is where OOXML puts it.

`titles: n` repeats the top `n` rows at the top of every printed page — the paper counterpart of `freeze`, which keeps them in view on screen. `0` clears it, as it does for a freeze. This one is written as a `_xlnm.Print_Titles` defined name in `xl/workbook.xml` rather than in the sheet part, and it is scoped to this worksheet alone, so each sheet of a workbook repeats its own rows.

`header` and `footer` are text printed at the top and the bottom of every page. Either is one **section**, printed on the left, or the three sections the format has — `{ left, center, right }`, each optional. A section is a string, or an array of **parts** in order:

- a string, printed as it is;
- `{ field: "page" }` or `{ field: "pages" }`, the page number and the page count, which the reader fills in as it paginates — `["Page ", { field: "page" }, " of ", { field: "pages" }]`;
- `{ text, bold, size }`, text in a look: `bold: true` and a font size from 1 to 99, each holding until a later part changes it, and starting plain at every section.

`firstHeader` and `firstFooter` are the first page's, where it differs from the rest. Naming either makes the first page different, and a first-page part you do not name prints nothing on that page — so a footer that should print everywhere goes in `footer`, and only what changes on page one goes in `firstFooter`.

**Text is text**: the format has a small language of `&`-codes for dates, file names and more, and this surface exposes none of it, so an ampersand in your text prints as an ampersand and a newline prints as a line break. A reader holds at most **255 characters** per header or footer as stored — the section codes, each field's code, each look's code and the doubled ampersands count — and a longer text throws rather than being cut mid-sentence by the reader. Nothing appears on screen; a header is print furniture, and the grid is unchanged.

### `sheet.place(id, { row, col, width, height })`

One floating picture anchored to the top-left of a 1-based cell, drawn at `width` × `height` CSS pixels at 96 dpi.

All indexing is **1-based**, uniformly — rows, columns and image anchors. The XML counts from zero and this package subtracts internally; that off-by-one is its business and appears nowhere in the API.

## Styles

A style is one object, and every part of it is optional:

```js
{
  font: { name, size, bold, italic, underline, strikethrough, color },
  fill: "#eeeeee",
  border: { top, right, bottom, left },   // each { style, color }, or absent
  alignment: { horizontal, vertical, wrapText },
  numberFormat: "#,##0.00",
}
```

Colours are `#rrggbb`, with the `#rgb` shorthand accepted. There is no alpha channel: a colour either paints or is absent.

A **fill is a colour**, not a pattern object — solid is the only pattern here, so there is nothing for a wrapper to distinguish. A border side's `style` is `"thin"`, `"dashed"` or `"dotted"`.

Identical fonts, fills, borders and formats collapse to one entry each, so a style repeated down a thousand rows costs one record.

## Guarantees, not options

There is nothing to configure. Each of these is a property of the writer rather than a parameter, because an option is a rule every caller has to hold in their head, and none of these is a choice a caller can make better:

- **Text is always interned** into the shared string table.
- **Style tables are always interned.**
- **Dates are the 1900 system**, converted in UTC. A reader's timezone never enters the conversion. This writer counts Excel's phantom 1900-02-29, so a date before March 1900 lands on the day it names in Excel. ECMA-376 defines the 1900 system with that fictitious day. LibreOffice Calc omits it and shows those dates one day early. No single serial satisfies both readers, so this writer follows the specification. A date from 1900-03-01 on reads the same in both, which covers every date a report is likely to hold.
- **A date with no format of its own** gets the short-date built-in, so it reads back as a day rather than as the number underneath it.
- **Media deduplicates by bytes.**
- **Every element is written in the Open XML SDK's child order.**
- **Created and modified are always `1970-01-01T00:00:00Z`** and cannot be overridden. Byte-identity is the headline promise, and an overridable clock would quietly redefine "the same input" as "the same input at the same wall time". Put a real timestamp in `description`.

## Determinism

The same calls produce byte-identical output **on the same runtime**, where a runtime means the engine and the zlib build inside it.

That qualifier is real rather than defensive. Deflate is not a single defined output: a conforming compressor may emit any valid stream, and engines differ. Chromium and Node happen to agree; Firefox and WebKit do not, and a Node patch release can carry a new zlib revision. So a digest of the finished archive pins a runtime as much as it pins a document.

If you want a fingerprint that survives all of that, hash the **parts** rather than the archive: unzip, then digest each part's bytes. Those are strings this package assembles, with no compressor anywhere in them.

## Errors

The writer throws native `TypeError` and `RangeError` at the boundary, with a message naming the offending argument. There is no custom error class and no error codes: any one of these is a mistake in the calling code rather than a condition to branch on, and zero dependencies means nothing to import in order to write a `catch`.

Leniency belongs where a document is authored. Strictness belongs here, because the job is never to emit a file a reader offers to repair.

One deliberate exception is silent rather than thrown: **text the format cannot carry**. Control characters other than tab, newline and carriage return are stripped, as are unpaired surrogates, and a literal `_xHHHH_` in your text is escaped so a reader does not decode it. Failing a whole document over one stray byte in a customer's name is the wrong trade.

## Content Security Policy

werkmap never turns a string into code. It runs under a policy with no `unsafe-eval` and no `unsafe-inline`, and its own suite proves it twice: the Node tests run with `--disallow-code-generation-from-strings`, and a Chromium page under `default-src 'none'; script-src 'self'` writes a workbook and reports any violation back.

The source is scanned for the two constructs on every run, so a regression fails the build rather than a deployment.

## Environments

Node.js 22 or newer, and browsers with `CompressionStream` — Chromium 103, Firefox 113, Safari 16.4. Workers included; there is no DOM access anywhere.

The package is plain JavaScript with JSDoc types. There is no build step: what npm installs is what runs.

## Contributing

Run the whole gate with `npm run check`. It formats, lints, checks for dead code and duplication, measures the bundle, runs the suite at 100% coverage, and drives the browser page.

## License

Apache-2.0
