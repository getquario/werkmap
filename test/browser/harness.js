// The strict-CSP browser test. What it proves is narrow and deliberate: the
// published file runs in a real browser under a policy that permits no
// string-to-code execution, and the archive it writes there is a valid one.
//
// It does not assert that the browser's bytes equal Node's. They were
// measured equal on Chromium, but this package promises byte-identity for the
// same input on the same runtime only, and a test asserting more would freeze
// an observation into a contract.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import http from "node:http";
import { chromium } from "playwright";

const html = await readFile(new URL("./index.html", import.meta.url), "utf8");

// The page loads its own modules and reaches nothing else. `script-src 'self'`
// with no `unsafe-*` of any kind is the whole point: a string-to-code path
// anywhere in the library would raise a violation the page reports back.
const CSP = [
  "default-src 'none'",
  "script-src 'self'",
  "connect-src 'none'",
  "img-src 'none'",
  "style-src 'none'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

// The bytes a browser gets are the bytes that ship. This suite exists to
// prove the published file runs under a strict policy, and a rewritten copy
// would prove it of something nobody installs — so nothing transforms a
// served module on the way out, and the assertion below fails if anything
// starts to.
const sources = new Map([
  ["/lib/index.js", new URL("../../lib/index.js", import.meta.url)],
  ["/fixture.js", new URL("../fixture.js", import.meta.url)],
]);

const files = new Map([
  ["/", ["text/html; charset=utf-8", html]],
  [
    "/browser.js",
    ["text/javascript; charset=utf-8", await readFile(new URL("./browser.js", import.meta.url))],
  ],
  ...(await Promise.all(
    [...sources].map(async ([path, source]) => [
      path,
      ["text/javascript; charset=utf-8", await readFile(source)],
    ]),
  )),
]);

const server = http.createServer((request, response) => {
  const file = files.get(new URL(request.url, "http://localhost").pathname);
  if (!file) {
    response.writeHead(404).end("Not found");
    return;
  }
  response
    .writeHead(200, {
      "Content-Type": file[0],
      "Content-Security-Policy": CSP,
      "X-Content-Type-Options": "nosniff",
    })
    .end(file[1]);
});

let browser;
try {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const { port } = server.address();
  browser = await chromium.launch();
  const page = await browser.newPage();
  const browserErrors = [];
  page.on("pageerror", (error) => browserErrors.push(error.message));
  page.on("console", (message) => {
    if (message.type() === "error") browserErrors.push(message.text());
  });

  for (const [path, source] of sources) {
    const served = await fetch(`http://127.0.0.1:${port}${path}`).then((r) => r.text());
    assert.equal(served, await readFile(source, "utf8"), `${path} is served rewritten`);
  }

  await page.goto(`http://127.0.0.1:${port}`);
  const done = page.locator('#result[data-status="passed"], #result[data-status="failed"]');
  await done.waitFor({ timeout: 10_000 });
  const status = await done.getAttribute("data-status");
  const message = await done.textContent();

  assert.equal(status, "passed", message);
  assert.deepEqual(browserErrors, []);
  console.log("Browser CSP test passed");
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
}
