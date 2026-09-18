# werkmap

A tiny spreadsheet writer for JavaScript. It writes an `.xlsx` file — rows, typed cells, styles, merged ranges, a frozen header, floating images — and nothing else. It does not read one. _Werkmap_ is Dutch for a workbook, which is the one thing this package makes.

Writing is the whole surface, so the package stays small enough to audit. The container is written against the ZIP specification over `CompressionStream`, every part is a string, and no value ever turns into code — so it runs unchanged under a Content Security Policy with no `unsafe-*` of any kind, which is what a spreadsheet export in the browser usually cannot do.

- **Byte-identical output.** Nothing consults a clock, a locale or a random source, and the ZIP epoch is pinned, so two renders of the same report are the same bytes and a build that caches by content hash keeps working.
- **Foreign readers accept it.** The suite loads every workbook it writes back through ExcelJS, an independent implementation — so the tests prove a real reader opens the file, not just that the bytes look plausible.
- **Zero dependencies.** 5.7 kB minified and brotlied, for the whole writer.
- **Strict CSP, including in the browser.** No `unsafe-eval` and no `unsafe-inline`. A Chromium page under `default-src 'none'; script-src 'self'` writes a workbook and reports any violation back, and the Node suite runs on `--disallow-code-generation-from-strings`.
- **Write-only, deliberately.** No reader, no formula engine, no chart support — see [Is werkmap the right tool?](#is-werkmap-the-right-tool) before you install it.
- **Hardened.** 53 tests at 100% branch coverage.

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
- You need column widths or row heights. A reader sizes columns from its own defaults.
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

### `sheet.row(cells) -> number`

`cells` is an array; the returned number is the row's 1-based position, which `merge`, `freeze` and `place` take.

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

### `sheet.print(setup)`

How this worksheet prints: `{ margin, size, orientation, fit, titles }`, every key optional. `margin` is all four page margins in points, `size` one of `letter`, `tabloid`, `legal`, `A3`, `A4`, `A5`, `orientation` either `portrait` or `landscape`, and `fit: true` scales the sheet to one page wide and as many pages tall as it takes.

A worksheet that never calls this carries **no print setup at all**, so a reader applies its own defaults rather than this writer's opinion. Calls merge, so two calls naming different keys both take effect. Print setup is per worksheet, which is where OOXML puts it.

`titles: n` repeats the top `n` rows at the top of every printed page — the paper counterpart of `freeze`, which keeps them in view on screen. `0` clears it, as it does for a freeze. This one is written as a `_xlnm.Print_Titles` defined name in `xl/workbook.xml` rather than in the sheet part, and it is scoped to this worksheet alone, so each sheet of a workbook repeats its own rows.

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
