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
 *
 * Every `fc.assert` / `fc.check` also takes its parameters from
 * `propertyConfig` (or runs through `assertProperty`, which does), so every
 * property gets the per-commit seed, the replay line and the nightly factor.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const SCANNED_ROOTS = ["packages", "scripts", "parity"] as const;
const BUDGET_HELPER = "propertyTestTimeout";
const RUN_COUNT_HELPER = "propertyConfig";
const ASSERT_HELPERS = new Set(["assertProperty", "assertPinnedProperty"]);
const FAST_CHECK_DRIVERS = new Set(["assert", "sample", "check"]);
const TEST_CALLEES = new Set(["test", "it"]);

setDefaultTimeout(120_000);

/** `fc.assert(...)`, `fc.sample(...)`, `fc.check(...)`. */
const callsFastCheck = (node: ts.CallExpression): boolean =>
  ts.isPropertyAccessExpression(node.expression) &&
  FAST_CHECK_DRIVERS.has(node.expression.name.text) &&
  ts.isIdentifier(node.expression.expression) &&
  node.expression.expression.text === "fc";

/** `assertProperty(...)`: `fc.assert` through `propertyConfig`. */
const callsAssertHelper = (node: ts.CallExpression): boolean =>
  ts.isIdentifier(node.expression) && ASSERT_HELPERS.has(node.expression.text);

const drivesFastCheck = (node: ts.CallExpression): boolean =>
  callsFastCheck(node) || callsAssertHelper(node);

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
  if (callsAssertHelper(call)) return true;
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

/**
 * `fc.assert(...)` / `fc.check(...)` calls whose parameters do not come from
 * `propertyConfig(...)`: they would miss the per-commit seed, the replay line
 * and the nightly factor.
 */
const bypassingSites = (file: string, sourceText: string): string[] => {
  const sourceFile = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
  const sites: string[] = [];
  const visit = (node: ts.Node): void => {
    if (
      ts.isCallExpression(node) &&
      callsFastCheck(node) &&
      ts.isPropertyAccessExpression(node.expression) &&
      node.expression.name.text !== "sample"
    ) {
      const params = node.arguments.at(1);
      const configured =
        params !== undefined &&
        ts.isCallExpression(params) &&
        ts.isIdentifier(params.expression) &&
        params.expression.text === RUN_COUNT_HELPER;
      if (!configured) {
        const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
        sites.push(`${file}:${String(line + 1)}`);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  return sites;
};

/** Nested configuration evaluates the pinned-seed guard before its assertion helper. */
const nestedConfigurationSites = (file: string, sourceText: string): string[] => {
  const sourceFile = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
  const sites: string[] = [];
  const visit = (node: ts.Node): void => {
    if (ts.isCallExpression(node) && callsAssertHelper(node)) {
      let params = node.arguments.at(1);
      while (params !== undefined && ts.isParenthesizedExpression(params))
        params = params.expression;
      if (
        params !== undefined &&
        ts.isCallExpression(params) &&
        ts.isIdentifier(params.expression) &&
        params.expression.text === RUN_COUNT_HELPER
      ) {
        const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
        sites.push(`${file}:${String(line + 1)}`);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
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

  test("reads which properties bypass propertyConfig", () => {
    expect(bypassingSites("probe.ts", "fc.assert(p, { numRuns: 4 });\nfc.check(p);")).toEqual([
      "probe.ts:1",
      "probe.ts:2",
    ]);
    expect(
      bypassingSites(
        "probe.ts",
        "fc.assert(p, propertyConfig());\nassertProperty(p, {});\nassertPinnedProperty(p, {});",
      ),
    ).toEqual([]);
  });

  test("every fc.assert and fc.check takes its parameters from propertyConfig", () => {
    const bypassing = testFiles().flatMap((file) => {
      const sourceText = ts.sys.readFile(path.join(REPO_ROOT, file));
      if (sourceText === undefined) panic(`Cannot read ${file}.`);
      return sourceText.includes("fc.") ? bypassingSites(file, sourceText) : [];
    });
    expect(bypassing).toEqual([]);
  });

  test("detects nested assertion configuration while accepting raw options and fast-check drivers", () => {
    for (const helper of ASSERT_HELPERS) {
      for (const config of [
        "propertyConfig()",
        "propertyConfig({ numRuns: 100 })",
        "(propertyConfig())",
      ]) {
        expect(nestedConfigurationSites("probe.ts", `${helper}(property, ${config});`)).toEqual([
          "probe.ts:1",
        ]);
      }
      expect(
        nestedConfigurationSites(
          "probe.ts",
          `${helper}(property);\n${helper}(property, {});\n${helper}(property, { numRuns: 100 });`,
        ),
      ).toEqual([]);
    }
    expect(
      nestedConfigurationSites(
        "probe.ts",
        "fc.assert(property, propertyConfig({ numRuns: 100 }));\nfc.check(property, propertyConfig());",
      ),
    ).toEqual([]);
    expect(
      nestedConfigurationSites(
        "probe.ts",
        '// assertProperty(property, propertyConfig());\nconst example = "assertPinnedProperty(property, propertyConfig())";',
      ),
    ).toEqual([]);
  });

  test("assertion helpers take raw options so pinned seeds replay before configuration", () => {
    const nested = testFiles().flatMap((file) => {
      const sourceText = ts.sys.readFile(path.join(REPO_ROOT, file));
      if (sourceText === undefined) panic(`Cannot read ${file}.`);
      return nestedConfigurationSites(file, sourceText);
    });
    expect(nested).toEqual([]);
  });

  test("every budget-requiring site declares propertyTestTimeout", () => {
    const undeclared = scanRepository()
      .filter(({ declaresBudget }) => !declaresBudget)
      .map(({ site }) => `${site} needs a stated budget with no ${BUDGET_HELPER}`);
    expect(undeclared).toEqual([]);
  });
});
