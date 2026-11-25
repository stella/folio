import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  censusUndescribedCanonicalCommands,
  checkCanonicalCommandBaseline,
} from "./lib/canonical-command-census";
import {
  canonicalCutoverDocs,
  inspectCanonicalSources,
} from "./lib/canonical-cutover";

const root = resolve(import.meta.dir, "..");
const commandsBaselineFile = "scripts/canonical-command-baseline.json";
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
const result = inspectCanonicalSources(
  files.map((file) => ({
    file,
    source: readFileSync(resolve(root, file), "utf8"),
  })),
);
const failures = [...result.failures];
const docs = canonicalCutoverDocs(result);
const write = process.argv.includes("--write");
const mergeBase = Bun.spawnSync(
  [
    "git",
    "merge-base",
    "HEAD",
    process.env["GITHUB_BASE_REF"]
      ? `origin/${process.env["GITHUB_BASE_REF"]}`
      : "origin/main",
  ],
  { cwd: root },
);
if (mergeBase.exitCode !== 0) throw new Error(mergeBase.stderr.toString());
const base = mergeBase.stdout.toString().trim();
const commandBaseline = Object.fromEntries(result.remainingCommands);
const parentPathsResult = Bun.spawnSync(
  [
    "git",
    "ls-tree",
    "-r",
    "--name-only",
    base,
    "packages/core/src/prosemirror",
  ],
  { cwd: root },
);
if (parentPathsResult.exitCode !== 0)
  throw new Error(parentPathsResult.stderr.toString());
const parentPaths = parentPathsResult.stdout
  .toString()
  .trim()
  .split("\n")
  .filter(
    (file) =>
      file.endsWith(".ts") &&
      !/\.test\.|\.typecheck\.|\/__tests__\/|\/__fixtures__\//.test(file),
  );
const parentBatch = Bun.spawnSync(["git", "cat-file", "--batch"], {
  cwd: root,
  stdin: new TextEncoder().encode(
    parentPaths.map((file) => `${base}:${file}\n`).join(""),
  ),
});
if (parentBatch.exitCode !== 0) throw new Error(parentBatch.stderr.toString());
const parentSources: { file: string; source: string }[] = [];
let cursor = 0;
for (const file of parentPaths) {
  const headerEnd = parentBatch.stdout.indexOf(10, cursor);
  if (headerEnd < 0) throw new Error("Incomplete git source snapshot header");
  const header = parentBatch.stdout.subarray(cursor, headerEnd).toString();
  const size = Number(header.split(" ").at(2));
  if (!header.includes(" blob ") || !Number.isSafeInteger(size) || size < 0)
    throw new Error(`Invalid git source snapshot: ${header}`);
  const start = headerEnd + 1;
  const end = start + size;
  if (end >= parentBatch.stdout.length || parentBatch.stdout[end] !== 10)
    throw new Error("Incomplete git source snapshot body");
  parentSources.push({
    file,
    source: parentBatch.stdout.subarray(start, end).toString(),
  });
  cursor = end + 1;
}
const priorCommands = censusUndescribedCanonicalCommands(parentSources);
const parentCommands = new Map(
  [...result.remainingCommands.keys()].map((gap) => [gap, priorCommands]),
);
for (const [gap, commands] of result.remainingCommands) {
  for (const command of checkCanonicalCommandBaseline(
    commands,
    parentCommands.get(gap) ?? [],
  ))
    failures.push(
      `${gap}: new command registration without a descriptor proof: ${command}`,
    );
}
if (
  !write &&
  readFileSync(resolve(root, commandsBaselineFile), "utf8") !==
    `${JSON.stringify(commandBaseline, null, 2)}\n`
)
  failures.push(
    "Canonical command baseline drifted; record decreases with --write",
  );
if (!write && readFileSync(resolve(root, docsFile), "utf8") !== docs)
  failures.push(
    "Canonical cutover docs drifted; regenerate with bun scripts/check-canonical-cutover.ts --write",
  );
if (failures.length > 0) {
  console.error(failures.join("\n"));
  process.exit(1);
}
if (write) {
  writeFileSync(resolve(root, docsFile), docs);
  writeFileSync(
    resolve(root, commandsBaselineFile),
    `${JSON.stringify(commandBaseline, null, 2)}\n`,
  );
}
console.log(
  `Canonical cutover: ${result.sites.size} capabilities; ${[...result.remainingCommands.values()].reduce((sum, commands) => sum + commands.length, 0)} remaining command registrations (base: ${[...parentCommands.values()].reduce((sum, commands) => sum + commands.length, 0)}).`,
);
