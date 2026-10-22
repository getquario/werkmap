// Whole-part equality, which the fragment matches elsewhere cannot reach: a
// constant no assertion names — a namespace URI, a content type, a
// relationship type — is still what decides whether a reader opens the file.
// This is not the stored byte golden the suite refuses: it compares the
// unzipped text, so no zlib revision reaches it. Regenerate with `npm run
// golden`.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { sample, xml } from "./helpers.js";

const dir = path.join(path.dirname(fileURLToPath(import.meta.url)), "golden");

test("every part matches the recorded document", async () => {
  for (const [name, text] of Object.entries(xml(await sample())))
    assert.equal(text, readFileSync(path.join(dir, name), "utf8"), name);
});
