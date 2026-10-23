/** Repository source guards that walk files or start processes state a scaled budget. */
import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";
import ts from "typescript";

import { scanTestTimeout } from "../test/scan-timeout";

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

const sourceGuardFiles = (): string[] => {
  const manifest = JSON.parse(
    ts.sys.readFile(path.join(REPO_ROOT, "package.json")) ?? "{}",
  ) as {
    scripts?: Record<string, string>;
  };
  return (manifest.scripts?.["test:source-guards"] ?? "")
    .split(/\s+/u)
    .filter((part) => /\.test\.(?:ts|mjs)$/u.test(part));
};

const callName = (expression: ts.Expression): string | undefined => {
  if (ts.isPropertyAccessExpression(expression)) return expression.name.text;
  if (
    ts.isElementAccessExpression(expression) &&
    ts.isStringLiteral(expression.argumentExpression)
  )
    return expression.argumentExpression.text;
  return ts.isIdentifier(expression) ? expression.text : undefined;
};

const callsAny = (
  sourceFile: ts.SourceFile,
  names: ReadonlySet<string>,
): boolean => {
  let found = false;
  const visit = (node: ts.Node): void => {
    if (found) return;
    if (ts.isCallExpression(node) || ts.isTaggedTemplateExpression(node)) {
      const name = callName(
        ts.isCallExpression(node) ? node.expression : node.tag,
      );
      if (name !== undefined && names.has(name)) {
        found = true;
        return;
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  return found;
};

const missingBudgets = (sources: ReadonlyMap<string, string>): string[] =>
  [...sources]
    .filter(([file, source]) => {
      const parsed = ts.createSourceFile(
        file,
        source,
        ts.ScriptTarget.Latest,
        true,
      );
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
      ts.sys.readFile(path.join(REPO_ROOT, file)) ?? "",
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
