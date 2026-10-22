// Rewrites test/golden/ from the current writer. Run `npm run golden` after an
// intentional change to the emitted XML and read the diff: that diff is the
// change, stated in the only terms a spreadsheet application sees.
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { sample, xml } from "./helpers.js";

const dir = path.join(path.dirname(fileURLToPath(import.meta.url)), "golden");

for (const [name, text] of Object.entries(xml(await sample()))) {
  const file = path.join(dir, name);
  await mkdir(path.dirname(file), { recursive: true });
  await writeFile(file, text);
  console.log("wrote", name);
}
