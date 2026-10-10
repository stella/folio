import { expect, test } from "bun:test";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const ROOT = path.resolve(import.meta.dir, "..");
const OWNERS = new Set([
  "packages/core/src/ai-edits/headless.ts",
  "packages/core/src/compare/import-style-closure.ts",
]);
const IMPORT_STEPS = new Set([
  "planTargetNumberingReferences",
  "stageTargetNumbering",
  "stageTargetStyles",
]);

const violations = (source: string): string[] => {
  if (![...IMPORT_STEPS].some((name) => source.includes(name))) return [];
  const file = ts.createSourceFile("fixture.ts", source, ts.ScriptTarget.Latest, true);
  const found: string[] = [];
  const memberName = (node: ts.Node) => {
    if (ts.isPropertyAccessExpression(node)) return node.name.text;
    if (ts.isElementAccessExpression(node) && ts.isStringLiteral(node.argumentExpression))
      return node.argumentExpression.text;
    if (ts.isBindingElement(node) && ts.isObjectBindingPattern(node.parent))
      return (node.propertyName ?? node.name).getText(file);
    return undefined;
  };
  const visit = (node: ts.Node): void => {
    const member = memberName(node);
    if (member !== undefined && IMPORT_STEPS.has(member)) found.push(node.getText(file));
    ts.forEachChild(node, visit);
  };
  visit(file);
  return found;
};

test("resource import steps cannot escape the shared sequence through aliases", () => {
  expect(violations("access.stageTargetStyles({})")).toHaveLength(1);
  expect(violations("const plan = access['planTargetNumberingReferences']; plan()")).toHaveLength(
    1,
  );
  expect(violations("const { stageTargetNumbering: stage } = access; stage()")).toHaveLength(1);
  expect(violations("importStyleClosureWithNumbering({})")).toEqual([]);
});

test("production resource imports use the shared numbering and style owner", () => {
  const files = execFileSync("git", ["ls-files", "packages/core/src"], { cwd: ROOT })
    .toString()
    .trim()
    .split("\n")
    .filter((file) => file.endsWith(".ts") && !file.endsWith(".test.ts") && !OWNERS.has(file));
  const found = files.flatMap((file) =>
    violations(readFileSync(path.join(ROOT, file), "utf8")).map((entry) => `${file}: ${entry}`),
  );
  expect(found).toEqual([]);
});
