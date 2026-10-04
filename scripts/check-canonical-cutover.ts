import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import {
  censusUndescribedCanonicalCommands,
  checkCanonicalCommandBaseline,
} from "./lib/canonical-command-census";
import {
  inspectCanonicalSources,
  checkCanonicalBaseline,
  canonicalCutoverDocs,
} from "./lib/canonical-cutover";

const readBaseline = (value: unknown): Record<string, number> => {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    throw new Error("Invalid canonical baseline");
  const counts: Record<string, number> = {};
  for (const [file, count] of Object.entries(value)) {
    if (typeof count !== "number" || !Number.isSafeInteger(count) || count <= 0)
      throw new Error(`Invalid canonical baseline count for ${file}`);
    counts[file] = count;
  }
  return counts;
};

const root = resolve(import.meta.dir, "..");
const baselineFile = "scripts/canonical-cutover-baseline.json";
const commandsBaselineFile = "scripts/canonical-command-baseline.json";
const docsFile = "docs/canonical-cutover.md";
const tracked = Bun.spawnSync(
  ["git", "ls-files", "packages/core/src", "packages/react/src", "packages/vue/src"],
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
  files.map((file) => ({ file, source: readFileSync(resolve(root, file), "utf8") })),
);
const failures = [...result.failures];
const docs = canonicalCutoverDocs(result);
const write = process.argv.includes("--write");
const baseline = write
  ? result.branches
  : readBaseline(JSON.parse(readFileSync(resolve(root, baselineFile), "utf8")));
if (!write) failures.push(...checkCanonicalBaseline(result.branches, baseline));
// Baseline edits cannot fund increases. On the introducing PR there is no parent baseline.
const mergeBase = Bun.spawnSync(
  [
    "git",
    "merge-base",
    "HEAD",
    process.env["GITHUB_BASE_REF"] ? `origin/${process.env["GITHUB_BASE_REF"]}` : "origin/main",
  ],
  { cwd: root },
);
if (mergeBase.exitCode !== 0) throw new Error(mergeBase.stderr.toString());
const base = mergeBase.stdout.toString().trim();
const exists = Bun.spawnSync(["git", "ls-tree", "--name-only", base, baselineFile], { cwd: root });
if (exists.exitCode !== 0) throw new Error(exists.stderr.toString());
if (exists.stdout.toString().trim() !== "") {
  const prior = Bun.spawnSync(["git", "show", `${base}:${baselineFile}`], { cwd: root });
  if (prior.exitCode !== 0) throw new Error(prior.stderr.toString());
  const parent = readBaseline(JSON.parse(prior.stdout.toString()));
  for (const [file, count] of Object.entries(baseline)) {
    const previous = parent[file] ?? 0;
    if (typeof count !== "number" || count > previous)
      failures.push(`${file}: committed baseline cannot increase above ${previous}`);
  }
}
const commandBaseline = Object.fromEntries(result.remainingCommands);
const parentPathsResult = Bun.spawnSync(
  ["git", "ls-tree", "-r", "--name-only", base, "packages/core/src/prosemirror"],
  { cwd: root },
);
if (parentPathsResult.exitCode !== 0) throw new Error(parentPathsResult.stderr.toString());
const parentPaths = parentPathsResult.stdout
  .toString()
  .trim()
  .split("\n")
  .filter(
    (file) =>
      file.endsWith(".ts") && !/\.test\.|\.typecheck\.|\/__tests__\/|\/__fixtures__\//.test(file),
  );
const parentBatch = Bun.spawnSync(["git", "cat-file", "--batch"], {
  cwd: root,
  stdin: new TextEncoder().encode(parentPaths.map((file) => `${base}:${file}\n`).join("")),
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
  parentSources.push({ file, source: parentBatch.stdout.subarray(start, end).toString() });
  cursor = end + 1;
}
const priorCommands = censusUndescribedCanonicalCommands(parentSources);
const parentCommands = new Map(
  [...result.remainingCommands.keys()].map((gap) => [gap, priorCommands]),
);
for (const [gap, commands] of result.remainingCommands) {
  for (const command of checkCanonicalCommandBaseline(commands, parentCommands.get(gap) ?? []))
    failures.push(`${gap}: new command registration without a descriptor proof: ${command}`);
}
if (
  !write &&
  readFileSync(resolve(root, commandsBaselineFile), "utf8") !==
    `${JSON.stringify(commandBaseline, null, 2)}\n`
)
  failures.push("Canonical command baseline drifted; record decreases with --write");
if (!write && readFileSync(resolve(root, docsFile), "utf8") !== docs)
  failures.push(
    "Canonical cutover docs drifted; regenerate with bun scripts/check-canonical-cutover.ts --write",
  );
if (failures.length > 0) {
  console.error(failures.join("\n"));
  process.exit(1);
}
if (write) {
  writeFileSync(resolve(root, baselineFile), `${JSON.stringify(result.branches, null, 2)}\n`);
  writeFileSync(resolve(root, docsFile), docs);
  writeFileSync(
    resolve(root, commandsBaselineFile),
    `${JSON.stringify(commandBaseline, null, 2)}\n`,
  );
}
console.log(
  `Canonical cutover: ${result.sites.size} capabilities; ${Object.values(result.branches).reduce((sum, count) => sum + count, 0)} session branches; ${[...result.remainingCommands.values()].reduce((sum, commands) => sum + commands.length, 0)} remaining command registrations (base: ${[...parentCommands.values()].reduce((sum, commands) => sum + commands.length, 0)}).`,
);
