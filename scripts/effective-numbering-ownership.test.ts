import { expect, test } from "bun:test";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const ROOT = path.resolve(import.meta.dir, "..");
const OWNER = "packages/core/src/prosemirror/numberingAttr.ts";
const MEMBERSHIP_READERS = new Set([
  "resolveListState",
  "resolveParagraphNumbering",
  "mergeParagraphNumbering",
  "paragraphNumberingReferenceId",
  "paragraphNumberingLevel",
]);
const EFFECTIVE_READERS = new Set([
  "effectiveParagraphNumbering",
  "effectiveParagraphNumberingReference",
]);

/** Follow local aliases so membership cannot accidentally read only the authored slot. */
const violations = (source: string): string[] => {
  // Files without either a source slot or a membership reader cannot cross
  // this boundary. Avoid constructing their syntax trees during discovery.
  if (!source.includes("numPr") || ![...MEMBERSHIP_READERS].some((name) => source.includes(name)))
    return [];
  const file = ts.createSourceFile("fixture.ts", source, ts.ScriptTarget.Latest, true);
  const scopeOf = (node: ts.Node): ts.Node => {
    let scope = node.parent;
    while (scope && !ts.isBlock(scope) && !ts.isSourceFile(scope) && !ts.isFunctionLike(scope))
      scope = scope.parent;
    return scope ?? file;
  };
  const bindings = new Map<ts.Node, Map<string, ts.Expression | undefined>>();
  const collect = (node: ts.Node): void => {
    if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name)) {
      const scope = scopeOf(node);
      const locals = bindings.get(scope) ?? new Map<string, ts.Expression | undefined>();
      locals.set(node.name.text, node.initializer);
      bindings.set(scope, locals);
    }
    if (
      ts.isVariableDeclaration(node) &&
      ts.isObjectBindingPattern(node.name) &&
      node.initializer
    ) {
      const scope = scopeOf(node);
      const locals = bindings.get(scope) ?? new Map<string, ts.Expression | undefined>();
      for (const element of node.name.elements) {
        if (!ts.isIdentifier(element.name) || element.dotDotDotToken) continue;
        const property = element.propertyName ?? element.name;
        if (!ts.isIdentifier(property)) continue;
        locals.set(
          element.name.text,
          ts.factory.createPropertyAccessExpression(node.initializer, property.text),
        );
      }
      bindings.set(scope, locals);
    }
    if (ts.isFunctionLike(node)) {
      const locals = bindings.get(node) ?? new Map<string, ts.Expression | undefined>();
      for (const parameter of node.parameters)
        if (ts.isIdentifier(parameter.name)) locals.set(parameter.name.text, undefined);
      bindings.set(node, locals);
    }
    ts.forEachChild(node, collect);
  };
  collect(file);
  const initializerOf = (node: ts.Identifier): ts.Expression | undefined => {
    let scope: ts.Node | undefined = node;
    while (scope) {
      const locals = bindings.get(scope);
      if (locals?.has(node.text)) return locals.get(node.text);
      scope = scope.parent;
    }
    return undefined;
  };
  const importNames = new Map<string, string>();
  const collectImports = (node: ts.Node): void => {
    if (ts.isImportSpecifier(node))
      importNames.set(node.name.text, (node.propertyName ?? node.name).text);
    ts.forEachChild(node, collectImports);
  };
  collectImports(file);
  const callableName = (node: ts.Expression, seen = new Set<string>()): string | undefined => {
    if (ts.isPropertyAccessExpression(node)) return node.name.text;
    if (!ts.isIdentifier(node)) return undefined;
    if (seen.has(node.text)) return undefined;
    const initializer = initializerOf(node);
    if (initializer) {
      const next = new Set(seen);
      next.add(node.text);
      return callableName(initializer, next);
    }
    return importNames.get(node.text) ?? node.text;
  };
  const isAttrs = (node: ts.Expression, seen = new Set<string>()): boolean => {
    if (ts.isIdentifier(node)) {
      if (/attrs$/iu.test(node.text)) return true;
      const initializer = initializerOf(node);
      if (initializer && !seen.has(node.text)) {
        const next = new Set(seen);
        next.add(node.text);
        return isAttrs(initializer, next);
      }
    }
    return (
      (ts.isPropertyAccessExpression(node) && node.name.text === "attrs") ||
      (ts.isCallExpression(node) &&
        ["expectParagraphAttrs", "mergeParagraphAttrs"].includes(
          callableName(node.expression) ?? "",
        ))
    );
  };
  const readsAuthoredSlot = (node: ts.Node, seen = new Set<string>()): boolean => {
    if (ts.isCallExpression(node) && EFFECTIVE_READERS.has(callableName(node.expression) ?? ""))
      return false;
    if (
      (ts.isPropertyAccessExpression(node) &&
        node.name.text === "numPr" &&
        isAttrs(node.expression)) ||
      (ts.isElementAccessExpression(node) &&
        ts.isStringLiteral(node.argumentExpression) &&
        node.argumentExpression.text === "numPr" &&
        isAttrs(node.expression))
    )
      return true;
    if (ts.isIdentifier(node) && !seen.has(node.text)) {
      const initializer = initializerOf(node);
      if (initializer) {
        const next = new Set(seen);
        next.add(node.text);
        return readsAuthoredSlot(initializer, next);
      }
    }
    return ts.forEachChild(node, (child) => readsAuthoredSlot(child, seen) || undefined) === true;
  };
  const found: string[] = [];
  const visit = (node: ts.Node): void => {
    if (
      ts.isCallExpression(node) &&
      MEMBERSHIP_READERS.has(callableName(node.expression) ?? "") &&
      node.arguments.some((argument) => readsAuthoredSlot(argument))
    )
      found.push(node.getText(file));
    ts.forEachChild(node, visit);
  };
  visit(file);
  return found;
};

test("membership readers require the effective numbering owner through local aliases", () => {
  expect(violations("resolveListState(map, attrs.numPr)")).toHaveLength(1);
  expect(
    violations("const direct = attrs['numPr']; resolveParagraphNumbering(direct)"),
  ).toHaveLength(1);
  expect(violations("resolveListState(map, effectiveParagraphNumbering(attrs))")).toEqual([]);
  expect(
    violations(
      "const effective = effectiveParagraphNumberingReference(attrs); resolveListState(map, effective)",
    ),
  ).toEqual([]);
  expect(violations("mergeParagraphNumbering(attrs.numPrFromStyle, attrs.numPr)")).toHaveLength(1);
  expect(
    violations(
      "const direct = attrs.numPr; function safe(direct) { resolveListState(map, direct); }",
    ),
  ).toEqual([]);
  expect(
    violations(
      "import { resolveListState as resolve } from './listState'; resolve(map, expectParagraphAttrs(node).numPr)",
    ),
  ).toHaveLength(1);
  expect(violations("const resolve = resolveListState; resolve(map, attrs.numPr)")).toHaveLength(1);
  expect(violations("const { numPr: direct } = attrs; resolveListState(map, direct)")).toHaveLength(
    1,
  );
});

test("production membership readers never consume the authored slot alone", () => {
  const files = execFileSync("git", ["ls-files", "packages/core/src"], { cwd: ROOT })
    .toString()
    .trim()
    .split("\n")
    .filter((file) => file.endsWith(".ts") && !file.endsWith(".test.ts") && file !== OWNER);
  const found = files.flatMap((file) =>
    violations(readFileSync(path.join(ROOT, file), "utf8")).map((entry) => `${file}: ${entry}`),
  );
  expect(found).toEqual([]);
});
