/**
 * A property test must state its own wall-clock budget once it can run long:
 * because it awaits per generated case, or because it raises `numRuns` above
 * fast-check's default.
 *
 * Bun kills a test at 5 s unless it says otherwise, and the nightly sweep sets
 * `PROPERTY_TEST_NUM_RUNS_FACTOR=10`, so every property runs ten times the
 * cases against whatever budget it declared. A synchronous predicate at the
 * stock run count finishes well inside 5 s, but the same predicate run ten
 * times as often at night no longer does once someone raises its `numRuns`;
 * an awaiting property can drift past 5 s on a single case, since every case
 * saves and reparses a package.
 *
 * So a site needs a budget when it awaits, or when it passes `numRuns`
 * explicitly, whether as a literal or through the repo's `propertyConfig`
 * helper (which always sets one). Either way the budget must come from
 * `propertyTestTimeout`, per test or once for the file: a plain number pins
 * PR CI and leaves the nightly run the same wall clock for ten times the work.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const SCANNED_ROOTS = ["packages", "scripts", "parity"] as const;
const BUDGET_HELPER = "propertyTestTimeout";
const RUN_COUNT_HELPER = "propertyConfig";
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

/**
 * `fc.assert(property, { numRuns: 40 })` or `fc.assert(property,
 * propertyConfig({ numRuns: 40 }))`. `propertyConfig` always establishes a
 * `numRuns` (100 by default), so any call through it counts as raising one
 * even when the literal it wraps omits the field.
 */
const passesNumRuns = (call: ts.CallExpression): boolean => {
  const params = call.arguments.at(1);
  if (params === undefined) return false;
  if (ts.isCallExpression(params)) {
    return ts.isIdentifier(params.expression) && params.expression.text === RUN_COUNT_HELPER;
  }
  return (
    ts.isObjectLiteralExpression(params) &&
    params.properties.some(
      (property) =>
        (ts.isPropertyAssignment(property) || ts.isShorthandPropertyAssignment(property)) &&
        ts.isIdentifier(property.name) &&
        property.name.text === "numRuns",
    )
  );
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

type BudgetRequiringSite = { site: string; declaresBudget: boolean };

const budgetRequiringSites = (file: string, sourceText: string): BudgetRequiringSite[] => {
  const sourceFile = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
  const fileBudget = declaresFileBudget(sourceText);
  const sites: BudgetRequiringSite[] = [];
  const visit = (node: ts.Node, enclosingTest: ts.CallExpression | null): void => {
    const nextTest = ts.isCallExpression(node) && isTestCall(node) ? node : enclosingTest;
    if (
      ts.isCallExpression(node) &&
      drivesFastCheck(node) &&
      nextTest !== null &&
      (isAsyncTest(nextTest) || passesNumRuns(node))
    ) {
      const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
      sites.push({
        site: `${file}:${String(line + 1)}`,
        declaresBudget: fileBudget || declaresOwnBudget(nextTest, sourceFile),
      });
    }
    ts.forEachChild(node, (child) => {
      visit(child, nextTest);
    });
  };
  visit(sourceFile, null);
  return sites;
};

const scanRepository = (): BudgetRequiringSite[] =>
  testFiles().flatMap((file) => {
    const sourceText = ts.sys.readFile(path.join(REPO_ROOT, file));
    if (sourceText === undefined) panic(`Cannot read ${file}.`);
    return sourceText.includes("fc.") ? budgetRequiringSites(file, sourceText) : [];
  });

describe("property test budgets", () => {
  test("reads where a budget-requiring site states its budget", () => {
    const body = "async () => { await fc.assert(p, propertyConfig({ numRuns: 40 })); }";
    expect(
      budgetRequiringSites("probe.ts", `test("x", ${body}, propertyTestTimeout(30_000));`),
    ).toEqual([{ site: "probe.ts:1", declaresBudget: true }]);
    expect(
      budgetRequiringSites(
        "probe.ts",
        `setDefaultTimeout(propertyTestTimeout(30_000));\ntest("x", ${body});`,
      ),
    ).toEqual([{ site: "probe.ts:2", declaresBudget: true }]);
    expect(budgetRequiringSites("probe.ts", `test("x", ${body});`)).toEqual([
      { site: "probe.ts:1", declaresBudget: false },
    ]);
    // A bare number pins PR CI and leaves the nightly sweep unscaled.
    expect(budgetRequiringSites("probe.ts", `test("x", ${body}, 30_000);`)).toEqual([
      { site: "probe.ts:1", declaresBudget: false },
    ]);
    // A synchronous property is a pure predicate at the default run count.
    expect(budgetRequiringSites("probe.ts", 'test("x", () => { fc.assert(p, {}); });')).toEqual([]);
    // Raising `numRuns` above the default multiplies the nightly wall clock
    // even without an `await`, so a sync site needs a budget too.
    expect(
      budgetRequiringSites("probe.ts", 'test("x", () => { fc.assert(p, { numRuns: 2000 }); });'),
    ).toEqual([{ site: "probe.ts:1", declaresBudget: false }]);
    expect(
      budgetRequiringSites(
        "probe.ts",
        'test("x", () => { fc.assert(p, propertyConfig({ numRuns: 2000 })); }, propertyTestTimeout(30_000));',
      ),
    ).toEqual([{ site: "probe.ts:1", declaresBudget: true }]);
    // `propertyConfig` sets `numRuns` even called bare, so it always counts.
    expect(
      budgetRequiringSites("probe.ts", 'test("x", () => { fc.assert(p, propertyConfig()); });'),
    ).toEqual([{ site: "probe.ts:1", declaresBudget: false }]);
  });

  test("scans the real sources, not the workspace symlinks", () => {
    const files = testFiles();
    expect(files.filter((file) => file.includes("/node_modules/"))).toEqual([]);
    expect(
      SCANNED_ROOTS.filter((root) => !files.some((file) => file.startsWith(`${root}/`))),
    ).toEqual([]);
    expect(scanRepository().length).toBeGreaterThan(0);
  });

  test("every budget-requiring site declares propertyTestTimeout", () => {
    const undeclared = scanRepository()
      .filter(({ declaresBudget }) => !declaresBudget)
      .map(({ site }) => `${site} needs a stated budget with no ${BUDGET_HELPER}`);
    expect(undeclared).toEqual([]);
  });
});
