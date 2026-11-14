import { expect, test } from "bun:test";
import { spawnSync } from "node:child_process";
import path from "node:path";

import baseline from "../.oxlint-plugins/folio-editor-commands.baseline.json";
import { subprocessTestTimeout } from "../test/subprocess";
import { resolveComparisonBase } from "./lib/comparison-base";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const BASELINE_PATH = ".oxlint-plugins/folio-editor-commands.baseline.json";

type Allowance = Readonly<Record<string, number>>;

/** The files whose allowance is new or larger than the base's. */
const grownAllowances = (current: Allowance, base: Allowance): string[] =>
  Object.entries(current)
    .filter(([file, count]) => count > (base[file] ?? 0))
    .map(([file]) => file);

test(
  "the raw command allowance never grows past the base's",
  () => {
    const baseSha = resolveComparisonBase({
      root: REPO_ROOT,
      requested: process.env["TOOLCHAIN_BASE"],
    });
    const listed = spawnSync(
      "git",
      ["ls-tree", "--name-only", baseSha, "--", BASELINE_PATH],
      { cwd: REPO_ROOT, encoding: "utf8" },
    );
    expect(listed.status).toBe(0);
    // The change that introduces the allowance has no base to compare with;
    // the lint rule holds every count to the sites that remain.
    if (listed.stdout.trim() === "") return;
    const shown = spawnSync("git", ["show", `${baseSha}:${BASELINE_PATH}`], {
      cwd: REPO_ROOT,
      encoding: "utf8",
    });
    expect(shown.status).toBe(0);
    const parsed: unknown = JSON.parse(shown.stdout);
    if (typeof parsed !== "object" || parsed === null)
      throw new TypeError("The base allowance must be a JSON object.");
    const base: Record<string, number> = {};
    for (const [file, count] of Object.entries(parsed))
      if (typeof count === "number") base[file] = count;
    expect(grownAllowances(baseline, base)).toEqual([]);
  },
  subprocessTestTimeout(2),
);

test("an allowance grows only by a new file or a larger count", () => {
  const base = { "packages/react/src/a.ts": 2 };
  expect(grownAllowances({ "packages/react/src/a.ts": 1 }, base)).toEqual([]);
  expect(grownAllowances({}, base)).toEqual([]);
  expect(grownAllowances({ "packages/react/src/a.ts": 3 }, base)).toEqual([
    "packages/react/src/a.ts",
  ]);
  expect(grownAllowances({ "packages/vue/src/b.ts": 1 }, base)).toEqual([
    "packages/vue/src/b.ts",
  ]);
});
