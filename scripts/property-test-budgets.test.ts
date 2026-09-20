/**
 * A property test that awaits per generated case must state its own budget.
 *
 * Bun kills a test at 5 s unless it says otherwise, and the nightly sweep sets
 * `PROPERTY_TEST_NUM_RUNS_FACTOR=10`, so every property runs ten times the
 * cases against whatever budget it declared. `numRuns` does not predict the
 * wall clock: the repo's four largest counts (1000-2000) each finish in under
 * a second, while the round-trip properties deliberately run 20-60 cases
 * because every case saves and reparses a package. The synchronous ones are
 * pure predicates; the ones that can drift past 5 s are the ones that await.
 *
 * So the line is `await`, not `numRuns`: an async test driving fast-check must
 * name a budget, per test or once for the file, and it must name it through
 * `propertyTestTimeout`. A plain number pins PR CI and leaves the nightly run
 * the same wall clock for ten times the work.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const SCANNED_ROOTS = ["packages", "scripts", "parity"] as const;
const BUDGET_HELPER = "propertyTestTimeout";
const FAST_CHECK_DRIVERS = new Set(["assert", "sample", "check"]);
const TEST_CALLEES = new Set(["test", "it"]);

setDefaultTimeout(120_000);

/** `fc.assert(...)`, `fc.sample(...)`, `fc.check(...)`. */
const drivesFastCheck = (node: ts.CallExpression): boolean =>
  ts.isPropertyAccessExpression(node.expression) &&
  FAST_CHECK_DRIVERS.has(node.expression.name.text) &&
  ts.isIdentifier(node.expression.expression) &&
  node.expression.expression.text === "fc";

/** `test(...)`, `it(...)`, and their `.each` / `.skip` / `.failing` forms. */
const isTestCall = (node: ts.CallExpression): boolean => {
  const callee = node.expression;
  if (ts.isIdentifier(callee)) return TEST_CALLEES.has(callee.text);
  if (!ts.isPropertyAccessExpression(callee)) return false;
  const owner = callee.expression;
  if (ts.isIdentifier(owner)) return TEST_CALLEES.has(owner.text);
  return (
    ts.isCallExpression(owner) &&
    ts.isPropertyAccessExpression(owner.expression) &&
    ts.isIdentifier(owner.expression.expression) &&
    TEST_CALLEES.has(owner.expression.expression.text)
  );
};

const isAsyncTest = (call: ts.CallExpression): boolean => {
  const body = call.arguments[1];
  if (body === undefined || (!ts.isArrowFunction(body) && !ts.isFunctionExpression(body))) {
    return false;
  }
  return body.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.AsyncKeyword) === true;
};

const declaresOwnBudget = (call: ts.CallExpression, sourceFile: ts.SourceFile): boolean => {
  const timeout = call.arguments.at(-1);
  if (timeout === undefined || call.arguments.length < 3) return false;
  return timeout.getText(sourceFile).includes(`${BUDGET_HELPER}(`);
};

const declaresFileBudget = (sourceText: string): boolean =>
  new RegExp(`setDefaultTimeout\\(\\s*${BUDGET_HELPER}\\(`).test(sourceText);

/**
 * `ts.sys.readDirectory` deduplicates by real path, so the workspace symlinks
 * under `packages/<name>/node_modules/@stll` shadow the packages they point at
 * and the real sources never get visited. Exclude them during the walk.
 */
const testFiles = (): string[] =>
  SCANNED_ROOTS.flatMap((root) =>
    ts.sys
      .readDirectory(path.join(REPO_ROOT, root), [".ts", ".tsx"], ["**/node_modules/**"])
      .map((absolute) => path.relative(REPO_ROOT, absolute).replaceAll("\\", "/"))
      .filter((file) => /\.test\.tsx?$/.test(file)),
  ).toSorted();

type AwaitingProperty = { site: string; declaresBudget: boolean };

const awaitingProperties = (file: string, sourceText: string): AwaitingProperty[] => {
  const sourceFile = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
  const fileBudget = declaresFileBudget(sourceText);
  const properties: AwaitingProperty[] = [];
  const visit = (node: ts.Node, enclosingTest: ts.CallExpression | null): void => {
    const nextTest = ts.isCallExpression(node) && isTestCall(node) ? node : enclosingTest;
    if (
      ts.isCallExpression(node) &&
      drivesFastCheck(node) &&
      nextTest !== null &&
      isAsyncTest(nextTest)
    ) {
      const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
      properties.push({
        site: `${file}:${String(line + 1)}`,
        declaresBudget: fileBudget || declaresOwnBudget(nextTest, sourceFile),
      });
    }
    ts.forEachChild(node, (child) => {
      visit(child, nextTest);
    });
  };
  visit(sourceFile, null);
  return properties;
};

const scanRepository = (): AwaitingProperty[] =>
  testFiles().flatMap((file) => {
    const sourceText = ts.sys.readFile(path.join(REPO_ROOT, file));
    if (sourceText === undefined) panic(`Cannot read ${file}.`);
    return sourceText.includes("fc.") ? awaitingProperties(file, sourceText) : [];
  });

describe("property test budgets", () => {
  test("reads where an awaiting property states its budget", () => {
    const body = "async () => { await fc.assert(p, propertyConfig({ numRuns: 40 })); }";
    expect(
      awaitingProperties("probe.ts", `test("x", ${body}, propertyTestTimeout(30_000));`),
    ).toEqual([{ site: "probe.ts:1", declaresBudget: true }]);
    expect(
      awaitingProperties(
        "probe.ts",
        `setDefaultTimeout(propertyTestTimeout(30_000));\ntest("x", ${body});`,
      ),
    ).toEqual([{ site: "probe.ts:2", declaresBudget: true }]);
    expect(awaitingProperties("probe.ts", `test("x", ${body});`)).toEqual([
      { site: "probe.ts:1", declaresBudget: false },
    ]);
    // A bare number pins PR CI and leaves the nightly sweep unscaled.
    expect(awaitingProperties("probe.ts", `test("x", ${body}, 30_000);`)).toEqual([
      { site: "probe.ts:1", declaresBudget: false },
    ]);
    // A synchronous property is a pure predicate; the 5 s default holds.
    expect(
      awaitingProperties("probe.ts", 'test("x", () => { fc.assert(p, { numRuns: 2000 }); });'),
    ).toEqual([]);
  });

  test("scans the real sources, not the workspace symlinks", () => {
    const files = testFiles();
    expect(files.filter((file) => file.includes("/node_modules/"))).toEqual([]);
    expect(
      SCANNED_ROOTS.filter((root) => !files.some((file) => file.startsWith(`${root}/`))),
    ).toEqual([]);
    expect(scanRepository().length).toBeGreaterThan(0);
  });

  test("every awaiting property declares propertyTestTimeout", () => {
    const undeclared = scanRepository()
      .filter(({ declaresBudget }) => !declaresBudget)
      .map(({ site }) => `${site} awaits per case with no ${BUDGET_HELPER}`);
    expect(undeclared).toEqual([]);
  });
});
