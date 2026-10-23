import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import ts from "typescript";

import { RUN_FORMATTING_MARK_NAMES } from "../../runFormattingMarkNames";

const MARK_EXTENSION_DIRECTORY = new URL(".", import.meta.url);

const canonicalWrapperBindings = (source: ts.SourceFile): Set<string> => {
  const bindings = new Set<string>();
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause) continue;
    const namedBindings = statement.importClause.namedBindings;
    if (!namedBindings || !ts.isNamedImports(namedBindings)) continue;
    for (const element of namedBindings.elements) {
      if (
        element.propertyName?.text === "withCanonicalCommand" ||
        element.name.text === "withCanonicalCommand"
      ) {
        bindings.add(element.name.text);
      }
    }
  }
  return bindings;
};

const runFormattingMarkNames = (source: ts.SourceFile): string[] => {
  const names: string[] = [];
  const visit = (node: ts.Node): void => {
    if (
      ts.isPropertyAssignment(node) &&
      node.name.getText(source) === "schemaMarkName" &&
      ts.isStringLiteral(node.initializer) &&
      RUN_FORMATTING_MARK_NAMES.has(node.initializer.text)
    ) {
      names.push(node.initializer.text);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return names;
};

const rawCanonicalWrapperCalls = (source: ts.SourceFile): string[] => {
  const bindings = canonicalWrapperBindings(source);
  const calls: string[] = [];
  const visit = (node: ts.Node): void => {
    if (
      ts.isCallExpression(node) &&
      ts.isIdentifier(node.expression) &&
      bindings.has(node.expression.text)
    ) {
      calls.push(node.expression.text);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return calls;
};

describe("run-formatting canonical boundary", () => {
  test("run mark extensions route canonical commands through the lock-aware helper", () => {
    const violations = readdirSync(MARK_EXTENSION_DIRECTORY)
      .filter((filename) => filename.endsWith("Extension.ts"))
      .flatMap((filename) => {
        const path = new URL(filename, MARK_EXTENSION_DIRECTORY);
        const source = ts.createSourceFile(
          filename,
          readFileSync(path, "utf8"),
          ts.ScriptTarget.Latest,
          true,
        );
        if (runFormattingMarkNames(source).length === 0) return [];
        const calls = rawCanonicalWrapperCalls(source);
        return calls.map((call) => `${filename}:${call}`);
      });
    expect(violations).toEqual([]);
  });
});
