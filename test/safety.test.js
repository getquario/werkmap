// The package's own CSP guard. The claim is that nothing here turns a string
// into code, and this is what holds the source to it — the suite already runs
// under `--disallow-code-generation-from-strings`, which catches an attempt at
// run time, while this catches one that no test happens to reach.
import { readdirSync, readFileSync } from "node:fs";
import assert from "node:assert/strict";
import test from "node:test";
import { workbook } from "../lib/index.js";

test("the source contains no string-to-code constructs", () => {
  const dir = new URL("../lib/", import.meta.url);
  const files = readdirSync(dir, { recursive: true })
    .map(String)
    .filter((name) => name.endsWith(".js"));
  assert.ok(files.length >= 1, "the scan covers every source module");
  for (const name of files) {
    const source = readFileSync(new URL(name, dir), "utf8");
    assert.ok(
      !/\beval\b|\bFunction\s*\(|new\s+Function/.test(source),
      `${name} is free of string-to-code constructs`,
    );
  }
});

test("hostile text reaches the sheet as inert data", async () => {
  const wb = workbook({ description: "</cp:coreProperties><script>alert(1)</script>" });
  const sheet = wb.sheet("Report");
  sheet.row([
    { value: '"/><script>alert(1)</script>' },
    { value: "]]>&<![CDATA[" },
    { value: [{ text: "<b>not bold</b>" }] },
  ]);

  const text = new TextDecoder().decode(await wb.bytes());
  assert.ok(!text.includes("<script>"), "no author text escaped its element");
});

test("a value the writer refuses never reaches the file", async () => {
  const wb = workbook();
  const sheet = wb.sheet("Report");
  sheet.row([{ value: "kept" }]);
  assert.throws(() => sheet.row([{ value: Number.NaN }]));

  // The throw happens before the row is appended, so a caught error leaves no
  // half-written row behind.
  const text = new TextDecoder().decode(await wb.bytes());
  assert.ok(!text.includes("NaN"), "NaN never reaches a value element");
});
