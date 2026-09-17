# werkmap

A write-only OOXML (`.xlsx`) writer. One entry point, `workbook`, a builder whose only
terminal is `bytes()`. Plain JS + JSDoc, **no build step** — `lib/` publishes verbatim.
Node 22+, ESM only, ES2024.

`README.md` is the normative description of the surface. A change to what a caller can
observe updates it in the same commit.

## Hard constraints

1. **Zero runtime dependencies.** Not a preference — it is why this package exists rather
   than a wrapper around an existing writer. A dependency here has to justify itself
   against that.
2. **No string-to-code, anywhere.** The suite runs under
   `--disallow-code-generation-from-strings`, a Chromium page runs under
   `default-src 'none'; script-src 'self'`, and `test/safety.test.js` scans `lib/` for the
   two constructs — so do not write the words naming them, even in a comment.
3. **Byte-identity for the same calls on the same runtime.** Nothing reads a clock, a
   locale, a random source or the environment. Dates are pinned, ZIP timestamps are pinned
   at the 1980 epoch, and every table is written in first-seen order. A change that makes
   output depend on anything else is a defect, not a trade.
4. **Open XML SDK child order, everywhere.** Free here, since the parts are assembled as
   strings in one pass, and it removes the ordering risk in readers whose tolerance nobody
   has measured.
5. **One file per entry point.** `lib/index.js` is the whole writer. Do not split it by
   concern; a second file arrives only with a second entry point.

## Where things are

- `lib/index.js` — ZIP container, XML helpers, value conversion, the style and
  shared-string tables, the sheet, the workbook. In that order, and the section rules mark
  the boundaries.
- `lib/index.d.ts` — hand-written, and the surface a consumer reads. It is not generated;
  keep it in step by hand.
- `test/helpers.js` — the two Node oracles: `book`/`sheet` read the bytes back through
  **exceljs**, an independent implementation, and `parts`/`xml` say what was actually
  written.
- `test/fixture.js` — the one fixture both halves share, imported by Node and **served to
  the browser page**, which is why it imports nothing.
- `test/browser/` — the page, and the harness that serves `lib/` unmodified under a strict
  policy.
- `bench/index.js` — `npm run bench`, the manual performance instrument, the same deal as in
  padvinder and sjabloon: no gate, no baseline, a table for a human. It times a row append
  and the terminal at 1k, 10k and 100k rows and prints the heap the builder holds, which is
  the number behind the README's "no streaming" line.

## What the suite may not assert

- **Not** that a browser's bytes equal Node's. They were measured equal on Chromium, but
  the promise is per-runtime; asserting more would freeze an observation into a contract.
- **Not** a stored byte golden. "Node 22" is not one zlib — patch releases carry new fork
  revisions — so a golden pins a runtime rather than a document. Byte-identity is tested by
  **self-comparison**: pack the same input twice and compare.

## Coverage

`test:unit` runs under `c8 --100`, and the bar is not negotiable: in a writer the untested
branch **is** the bug, because an unexercised style permutation emits XML nobody has ever
opened. No `c8 ignore` comments — a branch that cannot be reached should not exist. That is
also why `.fallowrc.jsonc` neutralises CRAP: it is the weaker form of a measure this
package already enforces exactly.

## Before you are done

`npm run check` — format, lint, dead code and duplication, size budget, the suite at 100%,
and the browser page. Commits are conventional; release-please cuts versions from them.
