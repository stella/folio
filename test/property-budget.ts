import { readFileSync } from "node:fs";
import ts from "typescript";

import { propertyTestTimeout } from "./property-timeout";

/**
 * The time budget a test states for itself, read from its source the way the
 * budget check (scripts/property-test-budgets.test.ts) reads it: the test
 * call's own timeout argument, else the file's `setDefaultTimeout`. A budget
 * written as `propertyTestTimeout(base)` is scaled by the run factor, exactly
 * as Bun receives it.
 */
export type StatedTestBudget = {
  /** `<absolute file>:<line of the test call>`, one per test definition. */
  key: string;
  /** The stated budget in milliseconds, or `undefined` when none is stated. */
  budgetMs: number | undefined;
};

type TestSpan = {
  startLine: number;
  endLine: number;
  /** A number, `"unknown"` when the argument cannot be evaluated, `undefined` when absent. */
  timeout: number | "unknown" | undefined;
};

type FileBudgets = {
  defaultTimeout: number | "unknown" | undefined;
  tests: readonly TestSpan[];
};

const TEST_CALLEES = new Set(["test", "it"]);

const rootName = (expression: ts.Expression): string | undefined => {
  if (ts.isIdentifier(expression)) return expression.text;
  if (ts.isPropertyAccessExpression(expression))
    return rootName(expression.expression);
  if (ts.isCallExpression(expression)) return rootName(expression.expression);
  return undefined;
};

const readBudgets = (file: string, text: string): FileBudgets => {
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
  const constants = new Map<string, ts.Expression[]>();
  const collect = (node: ts.Node): void => {
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.initializer !== undefined &&
      ts.isVariableDeclarationList(node.parent) &&
      (node.parent.flags & ts.NodeFlags.Const) !== 0
    ) {
      const existing = constants.get(node.name.text);
      if (existing) existing.push(node.initializer);
      else constants.set(node.name.text, [node.initializer]);
    }
    ts.forEachChild(node, collect);
  };
  collect(source);

  const evaluate = (
    expression: ts.Expression,
    resolving: ReadonlySet<string> = new Set(),
  ): number | undefined => {
    if (ts.isNumericLiteral(expression)) return Number(expression.text);
    if (ts.isParenthesizedExpression(expression))
      return evaluate(expression.expression, resolving);
    if (
      ts.isPrefixUnaryExpression(expression) &&
      expression.operator === ts.SyntaxKind.PlusToken
    )
      return evaluate(expression.operand, resolving);
    if (ts.isBinaryExpression(expression)) {
      const left = evaluate(expression.left, resolving);
      const right = evaluate(expression.right, resolving);
      if (left === undefined || right === undefined) return undefined;
      switch (expression.operatorToken.kind) {
        case ts.SyntaxKind.PlusToken:
          return left + right;
        case ts.SyntaxKind.AsteriskToken:
          return left * right;
        default:
          return undefined;
      }
    }
    if (ts.isIdentifier(expression)) {
      const definitions = constants.get(expression.text);
      const definition =
        definitions?.length === 1 ? definitions.at(0) : undefined;
      if (definition === undefined || resolving.has(expression.text))
        return undefined;
      return evaluate(definition, new Set([...resolving, expression.text]));
    }
    if (
      ts.isCallExpression(expression) &&
      ts.isIdentifier(expression.expression) &&
      expression.expression.text === "propertyTestTimeout" &&
      expression.arguments.length === 1
    ) {
      const base = evaluate(
        expression.arguments[0] as ts.Expression,
        resolving,
      );
      return base === undefined || !Number.isFinite(base) || base <= 0
        ? undefined
        : propertyTestTimeout(base);
    }
    if (ts.isObjectLiteralExpression(expression)) {
      const timeout = expression.properties.find(
        (property): property is ts.PropertyAssignment =>
          ts.isPropertyAssignment(property) &&
          ts.isIdentifier(property.name) &&
          property.name.text === "timeout",
      );
      return timeout === undefined
        ? undefined
        : evaluate(timeout.initializer, resolving);
    }
    return undefined;
  };
  const timeoutOf = (
    expression: ts.Expression | undefined,
  ): number | "unknown" | undefined => {
    if (expression === undefined) return undefined;
    const value = evaluate(expression);
    return value !== undefined && Number.isFinite(value) && value > 0
      ? value
      : "unknown";
  };
  const lineOf = (position: number): number =>
    source.getLineAndCharacterOfPosition(position).line + 1;

  let defaultTimeout: FileBudgets["defaultTimeout"];
  for (const statement of source.statements) {
    if (
      ts.isExpressionStatement(statement) &&
      ts.isCallExpression(statement.expression) &&
      ts.isIdentifier(statement.expression.expression) &&
      statement.expression.expression.text === "setDefaultTimeout"
    )
      defaultTimeout = timeoutOf(statement.expression.arguments.at(0));
  }
  const tests: TestSpan[] = [];
  const visit = (node: ts.Node): void => {
    if (
      ts.isCallExpression(node) &&
      TEST_CALLEES.has(rootName(node.expression) ?? "") &&
      node.arguments.length >= 2
    )
      tests.push({
        startLine: lineOf(node.getStart(source)),
        endLine: lineOf(node.getEnd()),
        timeout:
          node.arguments.length >= 3
            ? timeoutOf(node.arguments.at(-1))
            : undefined,
      });
    ts.forEachChild(node, visit);
  };
  visit(source);
  return { defaultTimeout, tests };
};

const files = new Map<string, FileBudgets | undefined>();

const budgetsOf = (absolute: string): FileBudgets | undefined => {
  if (!files.has(absolute)) {
    let budgets: FileBudgets | undefined;
    try {
      budgets = readBudgets(absolute, readFileSync(absolute, "utf8"));
    } catch {
      budgets = undefined;
    }
    files.set(absolute, budgets);
  }
  return files.get(absolute);
};

/**
 * The test enclosing the innermost of `frames` (absolute file, 1-based line)
 * that lies inside a test call, with the budget that test states; `undefined`
 * when no frame is inside a test. An unreadable budget counts as unstated.
 */
export const statedTestBudget = (
  frames: readonly { absolute: string; line: number }[],
): StatedTestBudget | undefined => {
  for (const { absolute, line } of frames) {
    if (!/\.tsx?$/.test(absolute) || absolute.includes("/node_modules/"))
      continue;
    const budgets = budgetsOf(absolute);
    if (budgets === undefined) continue;
    // Spans are in source order, so the last one holding the line is innermost.
    const enclosing = budgets.tests.findLast(
      (span) => span.startLine <= line && line <= span.endLine,
    );
    if (enclosing === undefined) continue;
    const stated = enclosing.timeout ?? budgets.defaultTimeout;
    return {
      key: `${absolute}:${String(enclosing.startLine)}`,
      budgetMs: typeof stated === "number" ? stated : undefined,
    };
  }
  return undefined;
};
