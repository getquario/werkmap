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
opened. No `c8 ignore` comments — a branch that cannot be reached should not exist.

**`fallow` is told, which is why the gate runs in two halves.** Its CRAP score is complexity
times untestedness, so a run that cannot see coverage estimates every function as untested,
`cc + cc²` puts the ceiling of 30 at cyclomatic **4**, and the gate enforces a number nobody
chose. So `fallow:health` takes `--coverage` and runs **after** the suite, in the job that owns
testing; `fallow:lint` is the dead-code and duplication half, which needs no coverage and stays
with the other linters. `npm run fallow` runs both, in that order, and `check` interleaves them
the same way CI does.

`fallow:health` **fails** when the coverage file is missing rather than falling back to an
estimate, which is what stops the gate quietly weakening. Run the suite first, or run
`npm run check`. A stale file is the other half of the same trap: coverage helps only where it
matches the current source, and fallow falls back to estimating when it does not.

## Code comments

A comment carries a _why_ the code cannot: a constraint, a deliberate deviation, a gotcha, a workaround. The code already shows the _how_, so the default is no comment.

- **Write for a reader who sees the file fresh.** The comment describes the code as it stands. What changed, and why it changed, goes in the commit message.
- **Keep the one fact a reader needs at that line.** An invariant the code cannot state ("the timeout stays below the poll interval; the host kills longer waits") or a sync obligation with another file ("mirror the list in `lib/index.d.ts`"). A comment that only restates a decision the code already reflects is deleted, even one that points at a doc.
- **The comment stands with every link removed.** Encode the substance; a link is a trailing breadcrumb, never the substance. Point at a maintained doc at a stable path (an ADR, `CONTEXT.md`, a README); a spec section number or a design doc is a point-in-time artifact that rots. When the why is a system-level narrative, it lives in that doc in full, and the comment keeps only the local detail.
- **Razor every comment you keep.** "Carries a real why" and "worded minimally" are separate checks. Cut the mechanism the code shows, where the value is consumed, the consequence of the consequence, the justification of the justification. A five-line block is suspect on sight; the razored answer is sometimes zero lines.
- **A public export gets a one-line summary.** A single clear line inside a body gets nothing.
- **A TODO is a marker.** It needs no issue ID, and it never stands in for work that is in scope.

## Before you are done

`npm run check` — format, lint, dead code and duplication, size budget, the suite at 100%,
and the browser page. Commits are conventional; release-please cuts versions from them.
