// Manual micro- and scaling benchmarks for werkmap. Run with `npm run bench`.
//
// Two costs are worth watching in a writer: what a row costs to append, and
// what the terminal costs to assemble, deflate and pack. The scaling table
// also prints how much heap the builder holds once the rows are in, because "does this fit in memory" is the question a large export
// asks first, and the README's "no streaming" line is only honest while the
// answer is measured rather than assumed.
import assert from "node:assert/strict";
import { workbook } from "../lib/index.js";
import { PNG } from "../test/fixture.js";

let sink = 0;

function consume(value) {
  sink += value instanceof Uint8Array ? value.length : typeof value === "number" ? 1 : 0;
}

function micro(name, fn) {
  for (let t = performance.now(); performance.now() - t < 50;) consume(fn());
  let best = 0;
  for (let sample = 0; sample < 5; sample++) {
    let ops = 0;
    const start = performance.now();
    let elapsed;
    do {
      for (let i = 0; i < 100; i++) consume(fn());
      ops += 100;
      elapsed = performance.now() - start;
    } while (elapsed < 100);
    best = Math.max(best, ops / (elapsed / 1e3));
  }
  console.log(name.padEnd(30), Math.round(best).toLocaleString().padStart(14), "ops/sec");
}

// The terminal is async, so its microbenchmark counts awaited calls rather
// than a tight loop: each sample runs as many as fit in 100 ms.
async function microAsync(name, fn) {
  for (let t = performance.now(); performance.now() - t < 50;) consume(await fn());
  let best = 0;
  for (let sample = 0; sample < 5; sample++) {
    let ops = 0;
    const start = performance.now();
    let elapsed;
    do {
      consume(await fn());
      ops += 1;
      elapsed = performance.now() - start;
    } while (elapsed < 100);
    best = Math.max(best, ops / (elapsed / 1e3));
  }
  console.log(name.padEnd(30), Math.round(best).toLocaleString().padStart(14), "ops/sec");
}

async function elapsed(name, fn) {
  consume(await fn());
  let best = Infinity;
  for (let sample = 0; sample < 3; sample++) {
    const start = performance.now();
    const result = await fn();
    const duration = performance.now() - start;
    consume(result);
    best = Math.min(best, duration);
  }
  console.log(name.padEnd(30), best.toFixed(3).padStart(14), "ms");
  return best;
}

const mb = (bytes) => (bytes / 1_048_576).toFixed(1).padStart(14) + " MB";

// Heap held right now, after a collection so the number is what is retained
// rather than what the collector has not got to yet. `--expose-gc` is in the
// script, so `gc` is always there. The package itself is not in this number:
// a Uint8Array's bytes live outside the JavaScript heap.
function heap() {
  globalThis.gc();
  return process.memoryUsage().heapUsed;
}

// A sales sheet: the report shape this writer exists for. Eight columns of
// mixed kinds, a bold frozen header merged over a title, a small pool of
// products and regions so the shared-string table gets exercised as a table
// rather than as a list, and a logo floating over the first row.
const PRODUCTS = Array.from({ length: 50 }, (_, i) => `Product ${i}`);
const REGIONS = ["North", "East", "South", "West"];
const HEADER = { font: { bold: true } };
const MONEY = { numberFormat: "#,##0.00" };

function cells(i) {
  return [
    { value: i + 1 },
    { value: PRODUCTS[i % PRODUCTS.length] },
    { value: REGIONS[i % REGIONS.length] },
    { value: (i % 17) + 1 },
    { value: ((i % 997) + 1) * 1.25, style: MONEY },
    { value: new Date(Date.UTC(2024, 0, 1 + (i % 365))) },
    { value: i % 3 === 0 },
    { value: `Order ${i + 1} for ${PRODUCTS[i % PRODUCTS.length]}` },
  ];
}

function build(size) {
  const wb = workbook({ title: "Sales", creator: "bench" });
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "Sales", style: HEADER }, null, null, null, null, null, null, null]);
  sheet.merge(1, 1, 8);
  sheet.place(logo, { row: 1, col: 8, width: 16, height: 16 });
  sheet.row(
    ["Id", "Product", "Region", "Qty", "Price", "Date", "Active", "Note"].map((value) => ({
      value,
      style: HEADER,
    })),
  );
  sheet.freeze(2);
  for (let i = 0; i < size; i++) sheet.row(cells(i));
  return wb;
}

// The one check the bench owes: the same calls yield the same bytes, and the
// small workbook is a real package rather than an empty one.
const first = await build(10).bytes();
assert.deepEqual(await build(10).bytes(), first);
assert.ok(first.length > 1_000);

console.log(`Node ${process.version} · ${process.platform} ${process.arch}`);
console.log("\nMicrobenchmarks (best of 5)");
// A sheet holds at most 1,048,576 rows and a fast machine appends more than
// that inside two microbenchmarks, so the sheet is swapped well before then.
let target = build(0).sheet("Rows");
const append = (row) => {
  const at = target.row(row);
  if (at === 500_000) target = build(0).sheet("Rows");
  return at;
};
const row = cells(1);
micro("row: 8 mixed cells", () => append(row));
micro("row: 8 mixed cells, built", () => append(cells(sink & 0xffff)));
const ten = build(10);
await microAsync("bytes: 10 rows", () => ten.bytes());

console.log("\nScaling (best of 3)");
for (const size of [1_000, 10_000, 100_000]) {
  console.log(`\n${size.toLocaleString()} rows`);
  const idle = heap();
  const wb = build(size);
  const builder = heap() - idle;
  await elapsed("rows", () => build(size));
  await elapsed("bytes()", () => wb.bytes());
  console.log("package".padEnd(30), mb((await wb.bytes()).length));
  console.log("heap held by the builder".padEnd(30), mb(builder));
}

if (sink < 0) console.log(sink);
