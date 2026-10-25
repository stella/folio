/** Repository source guards that walk files or start processes state a scaled budget. */
import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  type Node,
  parseSourceModule,
  type Program,
  staticMemberName,
  walkAst,
} from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const BUDGET_HELPERS = new Set(["propertyTestTimeout", "scanTestTimeout"]);
const WALK_METHODS = new Set([
  "readDirectory",
  "readdir",
  "readdirSync",
  "scan",
  "scanSync",
]);
// `$` is the shell template tag (`Bun.$` or an imported `$`).
const SUBPROCESS_METHODS = new Set([
  "$",
  "execFile",
  "execFileSync",
  "spawn",
  "spawnSync",
]);

setDefaultTimeout(scanTestTimeout(10_000));

const readSource = (file: string): string | undefined =>
  existsSync(file) ? readFileSync(file, "utf8") : undefined;

const sourceGuardFiles = (): string[] => {
  const manifest = JSON.parse(
    readSource(path.join(REPO_ROOT, "package.json")) ?? "{}",
  ) as {
    scripts?: Record<string, string>;
  };
  return (manifest.scripts?.["test:source-guards"] ?? "")
    .split(/\s+/u)
    .filter((part) => /\.test\.(?:ts|mjs)$/u.test(part));
};

const callName = (expression: Node): string | undefined =>
  staticMemberName(expression) ??
  (expression.type === "Identifier" ? expression.name : undefined);

const callsAny = (program: Program, names: ReadonlySet<string>): boolean => {
  let found = false;
  walkAst(program, (node) => {
    if (found) return false;
    if (
      node.type === "CallExpression" ||
      node.type === "TaggedTemplateExpression"
    ) {
      const name = callName(
        node.type === "CallExpression" ? node.callee : node.tag,
      );
      if (name !== undefined && names.has(name)) {
        found = true;
        return false;
      }
    }
    return true;
  });
  return found;
};

const missingBudgets = (sources: ReadonlyMap<string, string>): string[] =>
  [...sources]
    .filter(([file, source]) => {
      const parsed = parseSourceModule(file, source).program;
      return (
        (callsAny(parsed, WALK_METHODS) ||
          callsAny(parsed, SUBPROCESS_METHODS)) &&
        !callsAny(parsed, BUDGET_HELPERS)
      );
    })
    .map(([file]) => file)
    .toSorted();

test("repository-walking source guards declare a shared scaled budget", () => {
  const sources = new Map(
    sourceGuardFiles().map((file) => [
      file,
      readSource(path.join(REPO_ROOT, file)) ?? "",
    ]),
  );
  expect(missingBudgets(sources)).toEqual([]);
});

test("a repository walk without a budget is rejected", () => {
  const fixture = new Map([
    [
      "scripts/unbudgeted.test.ts",
      'import { test } from "bun:test";\ntest("walk", () => new Bun.Glob("**/*").scanSync());',
    ],
  ]);
  expect(missingBudgets(fixture)).toEqual(["scripts/unbudgeted.test.ts"]);
});

test("a shell subprocess without a budget is rejected", () => {
  const fixture = new Map([
    [
      "scripts/unbudgeted-shell.test.ts",
      'import { test } from "bun:test";\ntest("shell", () => Bun.$`bun scripts/check.ts`.quiet());',
    ],
  ]);
  expect(missingBudgets(fixture)).toEqual(["scripts/unbudgeted-shell.test.ts"]);
});
