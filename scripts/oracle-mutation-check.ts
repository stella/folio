#!/usr/bin/env bun
// A mutation survives when the packed consumer scenarios still pass after a
// deliberate product defect. The baseline tarballs are built once; each run
// receives its own repacked copy of folio-core, outside the checkout.

import { mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { repoRoot } from "./packaged-consumer-lib";

type Mutation = {
  id: string;
  bug: string;
  file: string;
  before: string;
  after: string;
  check: string;
  scenarios: string[];
  only: string;
  pr: boolean;
};

const MUTATIONS: Mutation[] = [
  {
    id: "table-cell-value",
    bug: "A populated inserted row loses its second cell value.",
    file: "ai-edits/table-row-column-mutations.js",
    before: 'const text = cellTexts[index] ?? "";',
    after: 'const text = index === 1 ? "" : cellTexts[index] ?? "";',
    check: "the requested-outcome oracle and table operation scenarios",
    scenarios: ["requested-outcome.test.ts", "operations.test.ts"],
    only: [
      "notices a payload the engine dropped or put on the wrong block",
      "tables / (direct|tracked-changes): every operation type applies or refuses, and the result saves",
      "tables / (direct|tracked-changes): every collision does what its applied operations asked",
    ].join("|"),
    pr: true,
  },
  {
    id: "inserted-last-word",
    bug: "An inserted paragraph loses its last word.",
    file: "ai-edits/apply.js",
    before: "for (const [paragraphIndex, text] of insertTexts.entries()) {",
    after:
      'for (const [paragraphIndex, originalText] of insertTexts.entries()) {\n\t\tconst text = originalText.replace(/\\s+\\S+$/u, "");',
    check: "the requested-outcome oracle and paragraph insertion scenarios",
    scenarios: ["requested-outcome.test.ts", "operations.test.ts"],
    only: "plain / direct: every (collision|operation type)",
    pr: true,
  },
  {
    id: "input-batch-order",
    bug: "Position-changing batch operations execute in input order.",
    file: "ai-edits/apply.js",
    before: "const executionOrder = executableResolved.toSorted((left, right) => {",
    after:
      "const executionOrder = executableResolved;\n\tconst unusedSortedOrder = executableResolved.toSorted((left, right) => {",
    check: "collision batch scenarios and fixed collision fuzz",
    scenarios: ["requested-outcome.test.ts", "operations.test.ts"],
    only: "(plain|tables) / direct: every collision does what its applied operations asked",
    pr: false,
  },
  {
    id: "nested-deletion-on-save",
    bug: "Saving an insertion omits its nested deletion mark.",
    file: "docx/serializer/paragraphSerializer.js",
    before:
      'case "mathEquation": return item.ommlXml;\n\t\t\tcase "insertion":\n\t\t\tcase "deletion":',
    after:
      'case "mathEquation": return item.ommlXml;\n\t\t\tcase "deletion": return "";\n\t\t\tcase "insertion":',
    check: "saved tracked-change follow-up scenarios",
    scenarios: ["requested-outcome.test.ts", "operations.test.ts"],
    only: "(replaceInPendingInsertion|replaceRangeInPendingInsertion|deletePendingInsertion) \\(tracked-changes\\)",
    pr: false,
  },
  {
    id: "false-applied-receipt",
    bug: "A skipped undefined-style operation reports applied.",
    file: "ai-edits/apply.js",
    before:
      'if (undefinedStyle !== void 0) {\n\t\t\tskipped.push({\n\t\t\t\tid: operation.id,\n\t\t\t\treason: "missingStyle",',
    after:
      'if (undefinedStyle !== void 0) {\n\t\t\tapplied.push({ id: operation.id });\n\t\t\tskipped.push({\n\t\t\t\tid: operation.id,\n\t\t\t\treason: "missingStyle",',
    check: "undefined-style receipt scenarios",
    scenarios: ["operations.test.ts", "requested-outcome.test.ts"],
    only: "insertAfterBlock with styleId NoSuchStyle is refused, not applied \\(direct\\)",
    pr: false,
  },
  {
    id: "stale-list-label",
    bug: "getContent uses a cached paragraph marker in place of the live counter.",
    file: "ai-edits/snapshot.js",
    before: "nextListLabel(expectParagraphAttrs(node))",
    after: "expectParagraphAttrs(node).listMarker ?? void 0",
    check: "list renumbering before save",
    scenarios: ["operations.test.ts"],
    only: "an item inserted into a list reads its own number",
    pr: true,
  },
  {
    id: "lost-footnote-reference",
    bug: "Saving a footnote reference drops its inline reference element.",
    file: "docx/serializer/runSerializer.js",
    before:
      'if (content.type === "footnoteRef") return `<w:footnoteReference w:id="${content.id}"${customMarkFollows}/>`;',
    after: 'if (content.type === "footnoteRef") return "";',
    check: "note reading-order scenario",
    scenarios: ["documents.test.ts"],
    only: "a note reference reads as its reading-order marker",
    pr: false,
  },
];

const PR_MODE = process.argv.includes("--pr");
const tarballsArg = process.argv.indexOf("--tarballs");
if (process.argv.some((arg) => arg.startsWith("--") && !["--pr", "--tarballs"].includes(arg))) {
  throw new Error("Usage: bun scripts/oracle-mutation-check.ts [--pr] [--tarballs DIR]");
}
const tarballsPath = tarballsArg === -1 ? undefined : process.argv.at(tarballsArg + 1);
if (tarballsArg !== -1 && tarballsPath === undefined) {
  throw new Error("--tarballs needs a directory");
}
const selected = MUTATIONS.filter((mutation) => !PR_MODE || mutation.pr);
const scratch = await mkdtemp(path.join(tmpdir(), "folio-oracle-mutations-"));
const baseline =
  tarballsPath === undefined ? path.join(scratch, "baseline") : path.resolve(tarballsPath);
const scenarioScript = path.join(repoRoot, "scripts", "consumer-scenarios.ts");
const fuzzEnv = {
  ...process.env,
  FOLIO_SCENARIO_SEED: "20260926",
  FOLIO_SCENARIO_FUZZ_RUNS: "2",
  FOLIO_SCENARIO_FUZZ_STEPS: "5",
  FOLIO_SCENARIO_COLLISION_RUNS: "2",
};

const run = async (command: string[], cwd = repoRoot) => {
  const child = Bun.spawn(command, { cwd, env: fuzzEnv, stdout: "pipe", stderr: "pipe" });
  const [stdout, stderr, exitCode] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ]);
  return { exitCode, output: `${stdout}\n${stderr}` };
};

const oneTarball = async (dir: string) => {
  const files = (await readdir(dir)).filter((file) => file.endsWith(".tgz"));
  const only = files.at(0);
  if (files.length !== 1 || only === undefined) {
    throw new Error(`Expected one tarball in ${dir}; found ${files.length}`);
  }
  return path.join(dir, only);
};

const patchExactlyOnce = (source: string, mutation: Mutation) => {
  const count = source.split(mutation.before).length - 1;
  if (count !== 1) {
    throw new Error(
      `${mutation.id}: expected one mutation seam in ${mutation.file}, found ${count}`,
    );
  }
  return source.replace(mutation.before, mutation.after);
};

const mutantTarballs = async (mutation: Mutation) => {
  const root = path.join(scratch, mutation.id);
  const extracted = path.join(root, "extracted");
  const core = path.join(root, "core");
  await mkdir(extracted, { recursive: true });
  await mkdir(core);
  for (const pkg of ["docx-core", "agents", "cli"]) {
    await symlink(path.join(baseline, pkg), path.join(root, pkg));
  }
  const original = await oneTarball(path.join(baseline, "core"));
  const unpack = await run(["tar", "-xzf", original, "-C", extracted]);
  if (unpack.exitCode !== 0)
    throw new Error(`${mutation.id}: tar extraction failed: ${unpack.output}`);
  const target = path.join(extracted, "package", "dist", mutation.file);
  await writeFile(target, patchExactlyOnce(await readFile(target, "utf8"), mutation));
  const archive = path.join(core, path.basename(original));
  const pack = await run(["tar", "-czf", archive, "-C", extracted, "package"]);
  if (pack.exitCode !== 0) throw new Error(`${mutation.id}: tar packaging failed: ${pack.output}`);
  return root;
};

const fuzzPattern = "fuzz run [01]|collision run [01]";

const scenarioArgs = (tarballDir: string, mutation: Mutation) => [
  "bun",
  scenarioScript,
  "--tarballs",
  tarballDir,
  "--only",
  `${mutation.only}|${fuzzPattern}`,
  "--",
  ...mutation.scenarios,
  "fuzz.test.ts",
];

try {
  if (tarballsArg === -1) {
    await mkdir(baseline);
    const pack = await run(["bun", scenarioScript, "--pack-only", baseline]);
    if (pack.exitCode !== 0) throw new Error(`Baseline package build failed:\n${pack.output}`);
  }
  // Cover every file the selected mutants exercise before interpreting a red run.
  const files = [...new Set(selected.flatMap((mutation) => mutation.scenarios))];
  const healthy = await run([
    "bun",
    scenarioScript,
    "--tarballs",
    baseline,
    "--only",
    `${selected.map((mutation) => mutation.only).join("|")}|${fuzzPattern}`,
    "--",
    ...files,
    "fuzz.test.ts",
  ]);
  if (healthy.exitCode !== 0) {
    throw new Error(`Unmutated consumer scenarios failed:\n${healthy.output}`);
  }
  let survivors = 0;
  for (const mutation of selected) {
    console.log(`→ ${mutation.id}: ${mutation.bug}`);
    const tarballDir = await mutantTarballs(mutation);
    const result = await run(scenarioArgs(tarballDir, mutation));
    const failures = [
      ...new Set(
        result.output
          .split("\n")
          .filter((line) => /^\s*(?:✖|not ok) /u.test(line) && !line.includes("failing tests:"))
          .map((line) => line.trim()),
      ),
    ];
    if (result.exitCode === 0) {
      console.error(`SURVIVED ${mutation.id}: expected ${mutation.check} to fail.`);
      survivors += 1;
    } else if (failures.length === 0) {
      throw new Error(`${mutation.id}: suite exited before a test failed:\n${result.output}`);
    } else {
      console.log(`CAUGHT ${mutation.id}: ${failures.join("; ")}`);
    }
  }
  if (survivors > 0) {
    throw new Error(
      `${survivors} of ${selected.length} mutations survived the consumer scenarios.`,
    );
  }
  console.log(`All ${selected.length} mutations were caught.`);
} finally {
  await rm(scratch, { recursive: true, force: true });
}
