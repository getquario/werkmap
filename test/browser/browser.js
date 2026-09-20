import { workbook } from "/lib/index.js";
import { PNG } from "/fixture.js";

const result = document.querySelector("#result");
const violations = [];

document.addEventListener("securitypolicyviolation", (event) => {
  violations.push(`${event.violatedDirective}: ${event.blockedURI}`);
});

const assert = (value, message) => {
  if (!value) throw Error(message);
};

const build = async () => {
  const wb = workbook({ title: "Sales", description: "Written in a browser" });
  const logo = wb.image(PNG, "png");
  const sheet = wb.sheet("Report");
  const head = sheet.row([
    { value: "Sales", style: { font: { bold: true, size: 16 }, fill: "#eeeeee" } },
    { value: null },
  ]);
  sheet.merge(head, 1, 2);
  sheet.row([
    { value: "<script>globalThis.pwned = true</script>" },
    { value: 1050, style: { numberFormat: "#,##0.00" } },
  ]);
  sheet.row([{ value: new Date(Date.UTC(2024, 1, 29)) }, { value: true }]);
  sheet.row(
    [[{ text: "Total " }, { text: "1050", font: { bold: true } }]].map((v) => ({ value: v })),
  );
  sheet.freeze(1);
  sheet.place(logo, { row: 1, width: 120, height: 40 });
  sheet.link(2, 1, { url: "https://example.test/mouse" });
  sheet.link(3, 1, { location: "'Report'!A1" });
  return wb.bytes();
};

try {
  globalThis.pwned = false;

  const bytes = await build();
  assert(bytes instanceof Uint8Array, "the writer did not produce bytes");
  assert(bytes[0] === 0x50 && bytes[1] === 0x4b, "the package is missing its zip header");

  // The promise is byte-identity for the same input on the same runtime, and
  // this is the runtime the claim is about here.
  const again = await build();
  assert(bytes.length === again.length, "two packs of the same input differ in length");
  assert(
    bytes.every((byte, at) => byte === again[at]),
    "two packs of the same input differ in content",
  );

  // Author text reaches the parts as data. The page asserts it here rather
  // than trusting the Node suite, because a browser is where an escaping
  // mistake would turn into something that runs.
  const text = new TextDecoder().decode(bytes);
  assert(!text.includes("<script>"), "author text escaped its element");
  assert(globalThis.pwned === false, "something in the written document ran");

  // `CompressionStream` is the one platform API the writer reaches for, so a
  // browser that packs a plausible archive is the thing worth proving.
  assert(bytes.length > 2000, `the package is implausibly small: ${bytes.length} bytes`);

  await new Promise((resolve) => setTimeout(resolve, 0));
  assert(violations.length === 0, `CSP violation: ${violations.join(", ")}`);
  result.dataset.status = "passed";
  result.textContent = "passed";
} catch (error) {
  result.dataset.status = "failed";
  result.textContent = error.stack || String(error);
}
