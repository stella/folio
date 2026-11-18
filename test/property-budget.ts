import { readFileSync } from "node:fs";

import {
  createLineIndex,
  type Node,
  parseSourceModule,
  walkAst,
} from "../scripts/lib/oxc-ast";
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

const rootName = (expression: Node): string | undefined => {
  if (expression.type === "Identifier") return expression.name;
  if (expression.type === "MemberExpression" && !expression.computed)
    return rootName(expression.object);
  if (expression.type === "CallExpression") return rootName(expression.callee);
  if (expression.type === "ChainExpression")
    return rootName(expression.expression);
  return undefined;
};

const readBudgets = (file: string, text: string): FileBudgets => {
  const { program } = parseSourceModule(file, text);
  const constants = new Map<string, Node[]>();
  walkAst(program, (node, ancestors) => {
    const list = ancestors.at(-1);
    if (
      node.type === "VariableDeclarator" &&
      node.id.type === "Identifier" &&
      node.init &&
      list?.type === "VariableDeclaration" &&
      // `await using` carries TypeScript's const flag too.
      (list.kind === "const" || list.kind === "await using")
    ) {
      const existing = constants.get(node.id.name);
      if (existing) existing.push(node.init);
      else constants.set(node.id.name, [node.init]);
    }
  });

  const evaluate = (
    expression: Node,
    resolving: ReadonlySet<string> = new Set(),
  ): number | undefined => {
    if (expression.type === "Literal") {
      const { value } = expression;
      return typeof value === "number" ? value : undefined;
    }
    if (expression.type === "ParenthesizedExpression")
      return evaluate(expression.expression, resolving);
    if (expression.type === "UnaryExpression" && expression.operator === "+")
      return evaluate(expression.argument, resolving);
    if (expression.type === "BinaryExpression") {
      const left = evaluate(expression.left, resolving);
      const right = evaluate(expression.right, resolving);
      if (left === undefined || right === undefined) return undefined;
      switch (expression.operator) {
        case "+":
          return left + right;
        case "*":
          return left * right;
        default:
          return undefined;
      }
    }
    if (expression.type === "Identifier") {
      const definitions = constants.get(expression.name);
      const definition =
        definitions?.length === 1 ? definitions.at(0) : undefined;
      if (definition === undefined || resolving.has(expression.name))
        return undefined;
      return evaluate(definition, new Set([...resolving, expression.name]));
    }
    if (
      expression.type === "CallExpression" &&
      expression.callee.type === "Identifier" &&
      expression.callee.name === "propertyTestTimeout" &&
      expression.arguments.length === 1
    ) {
      const argument = expression.arguments.at(0);
      const base =
        argument === undefined ? undefined : evaluate(argument, resolving);
      return base === undefined || !Number.isFinite(base) || base <= 0
        ? undefined
        : propertyTestTimeout(base);
    }
    if (expression.type === "ObjectExpression") {
      const timeout = expression.properties.find(
        (property) =>
          property.type === "Property" &&
          property.kind === "init" &&
          !property.method &&
          !property.shorthand &&
          !property.computed &&
          property.key.type === "Identifier" &&
          property.key.name === "timeout",
      );
      return timeout?.type === "Property"
        ? evaluate(timeout.value, resolving)
        : undefined;
    }
    return undefined;
  };
  const timeoutOf = (
    expression: Node | undefined,
  ): number | "unknown" | undefined => {
    if (expression === undefined) return undefined;
    const value = evaluate(expression);
    return value !== undefined && Number.isFinite(value) && value > 0
      ? value
      : "unknown";
  };
  const lineIndex = createLineIndex(text);
  const lineOf = (offset: number): number => lineIndex(offset).line + 1;

  let defaultTimeout: FileBudgets["defaultTimeout"];
  for (const statement of program.body) {
    if (
      statement.type === "ExpressionStatement" &&
      statement.expression.type === "CallExpression" &&
      statement.expression.callee.type === "Identifier" &&
      statement.expression.callee.name === "setDefaultTimeout"
    )
      defaultTimeout = timeoutOf(statement.expression.arguments.at(0));
  }
  const tests: TestSpan[] = [];
  walkAst(program, (node) => {
    if (
      node.type === "CallExpression" &&
      TEST_CALLEES.has(rootName(node.callee) ?? "") &&
      node.arguments.length >= 2
    )
      tests.push({
        startLine: lineOf(node.start),
        endLine: lineOf(node.end),
        timeout:
          node.arguments.length >= 3
            ? timeoutOf(node.arguments.at(-1))
            : undefined,
      });
  });
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
