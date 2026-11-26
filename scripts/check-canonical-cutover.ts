import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  censusUndescribedCanonicalCommands,
  undescribedCanonicalCommandFailures,
} from "./lib/canonical-command-census";
import {
  canonicalCutoverDocs,
  inspectCanonicalSources,
} from "./lib/canonical-cutover";

const root = resolve(import.meta.dir, "..");
const docsFile = "docs/canonical-cutover.md";
const tracked = Bun.spawnSync(
  [
    "git",
    "ls-files",
    "packages/core/src",
    "packages/react/src",
    "packages/vue/src",
  ],
  { cwd: root },
);
if (tracked.exitCode !== 0) throw new Error(tracked.stderr.toString());
const files = tracked.stdout
  .toString()
  .trim()
  .split("\n")
  .filter(
    (file) =>
      /\.(ts|tsx|vue)$/.test(file) &&
      !/\.test\.|\.typecheck\.|\/__tests__\/|\/__fixtures__\//.test(file),
  );
const sources = files.map((file) => ({
  file,
  source: readFileSync(resolve(root, file), "utf8"),
}));
const result = inspectCanonicalSources(sources);
const undescribed = censusUndescribedCanonicalCommands(
  sources.filter(
    ({ file }) =>
      file.startsWith("packages/core/src/prosemirror/") && file.endsWith(".ts"),
  ),
);
const failures = [
  ...result.failures,
  ...undescribedCanonicalCommandFailures(undescribed),
];
const docs = canonicalCutoverDocs(result);
const write = process.argv.includes("--write");
if (!write && readFileSync(resolve(root, docsFile), "utf8") !== docs)
  failures.push(
    "Canonical cutover docs drifted; regenerate with bun scripts/check-canonical-cutover.ts --write",
  );
if (failures.length > 0) {
  console.error(failures.join("\n"));
  process.exit(1);
}
if (write) writeFileSync(resolve(root, docsFile), docs);
console.log(`Canonical cutover: ${result.sites.size} capabilities.`);
