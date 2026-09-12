import { describe, expect, test } from "bun:test";
import { existsSync, readFileSync, realpathSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const { dirname, extname, relative, resolve } = path;

const COMPARE_DIRECTORY = realpathSync(resolve(import.meta.dir));
const NEUTRAL_ROOTS = ["content.ts", "content-types.ts"] as const;
const ALLOWED_EXTERNAL_IMPORTS = new Set(["better-result"]);
const BANNED_GLOBALS = new Set([
  "Atomics",
  "Bun",
  "console",
  "Date",
  "Deno",
  "eval",
  "FinalizationRegistry",
  "Function",
  "global",
  "globalThis",
  "indexedDB",
  "Intl",
  "localStorage",
  "Temporal",
  "crypto",
  "document",
  "fetch",
  "HTMLElement",
  "navigator",
  "performance",
  "process",
  "requestAnimationFrame",
  "self",
  "sessionStorage",
  "setInterval",
  "setTimeout",
  "SharedArrayBuffer",
  "WebSocket",
  "window",
  "Worker",
  "WeakRef",
  "XMLHttpRequest",
]);

type NeutralBoundaryViolation = {
  readonly file: string;
  readonly line: number;
  readonly reason: string;
};

type AnalyzedSource = {
  readonly staticSpecifiers: readonly string[];
  readonly violations: readonly NeutralBoundaryViolation[];
};

const relativeModuleName = (file: string): string => relative(COMPARE_DIRECTORY, file);

const violation = (
  sourceFile: ts.SourceFile,
  node: ts.Node,
  reason: string,
): NeutralBoundaryViolation => ({
  file: relativeModuleName(sourceFile.fileName),
  line: sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile)).line + 1,
  reason,
});

const isDeclarationName = (node: ts.Identifier): boolean => {
  const { parent } = node;
  return (
    ((ts.isVariableDeclaration(parent) ||
      ts.isParameter(parent) ||
      ts.isFunctionDeclaration(parent) ||
      ts.isClassDeclaration(parent) ||
      ts.isInterfaceDeclaration(parent) ||
      ts.isTypeAliasDeclaration(parent) ||
      ts.isEnumDeclaration(parent) ||
      ts.isBindingElement(parent) ||
      ts.isImportClause(parent) ||
      ts.isImportSpecifier(parent) ||
      ts.isNamespaceImport(parent) ||
      ts.isImportEqualsDeclaration(parent)) &&
      parent.name === node) ||
    ((ts.isPropertyDeclaration(parent) ||
      ts.isPropertySignature(parent) ||
      ts.isMethodDeclaration(parent) ||
      ts.isMethodSignature(parent) ||
      ts.isGetAccessorDeclaration(parent) ||
      ts.isSetAccessorDeclaration(parent)) &&
      parent.name === node) ||
    (ts.isPropertyAccessExpression(parent) && parent.name === node) ||
    (ts.isPropertyAssignment(parent) && parent.name === node) ||
    (ts.isLabeledStatement(parent) && parent.label === node)
  );
};

const isTypePosition = (node: ts.Identifier): boolean => {
  let current: ts.Node = node;
  while (current.parent) {
    const parent = current.parent;
    if (ts.isTypeNode(parent)) return true;
    if (
      ts.isExpressionWithTypeArguments(parent) ||
      ts.isHeritageClause(parent) ||
      ts.isTypeParameterDeclaration(parent)
    ) {
      current = parent;
      continue;
    }
    break;
  }
  return false;
};

const isRuntimeGlobalReference = (node: ts.Identifier): boolean =>
  BANNED_GLOBALS.has(node.text) && !isDeclarationName(node) && !isTypePosition(node);

const analyzeSource = (fileName: string, source: string): AnalyzedSource => {
  const sourceFile = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const staticSpecifiers: string[] = [];
  const violations: NeutralBoundaryViolation[] = [];

  const visit = (node: ts.Node): void => {
    if (
      (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) &&
      node.moduleSpecifier !== undefined &&
      ts.isStringLiteral(node.moduleSpecifier)
    ) {
      staticSpecifiers.push(node.moduleSpecifier.text);
    }
    if (ts.isImportEqualsDeclaration(node) && ts.isExternalModuleReference(node.moduleReference)) {
      violations.push(violation(sourceFile, node, "CommonJS import assignment is not allowed"));
    }
    if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
      violations.push(violation(sourceFile, node, "dynamic import is not allowed"));
    }
    if (
      ts.isCallExpression(node) &&
      ts.isIdentifier(node.expression) &&
      node.expression.text === "require"
    ) {
      violations.push(violation(sourceFile, node, "CommonJS require is not allowed"));
    }
    if (
      ts.isPropertyAccessExpression(node) &&
      ts.isIdentifier(node.expression) &&
      node.expression.text === "Math" &&
      node.name.text === "random"
    ) {
      violations.push(violation(sourceFile, node, "Math.random is not allowed"));
    }
    if (
      ts.isElementAccessExpression(node) &&
      ts.isIdentifier(node.expression) &&
      node.expression.text === "Math" &&
      ts.isStringLiteral(node.argumentExpression) &&
      node.argumentExpression.text === "random"
    ) {
      violations.push(violation(sourceFile, node, "Math.random is not allowed"));
    }
    if (
      ts.isPropertyAccessExpression(node) &&
      (node.name.text === "localeCompare" || node.name.text.startsWith("toLocale"))
    ) {
      violations.push(violation(sourceFile, node, "ambient-locale operations are not allowed"));
    }
    if (ts.isIdentifier(node) && isRuntimeGlobalReference(node)) {
      violations.push(
        violation(sourceFile, node, `runtime global ${JSON.stringify(node.text)} is not allowed`),
      );
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  return { staticSpecifiers, violations };
};

const resolveRelativeModule = (importer: string, specifier: string): string | null => {
  const unresolved = resolve(dirname(importer), specifier);
  const extension = extname(unresolved);
  const candidates = extension
    ? [unresolved]
    : [`${unresolved}.ts`, `${unresolved}.tsx`, resolve(unresolved, "index.ts")];
  const candidate = candidates.find((current) => existsSync(current));
  return candidate === undefined ? null : realpathSync(candidate);
};

type NeutralModuleGraph = {
  readonly modules: readonly string[];
  readonly violations: readonly NeutralBoundaryViolation[];
};

const neutralModuleGraph = (): NeutralModuleGraph => {
  const pending = NEUTRAL_ROOTS.map((root) => resolve(COMPARE_DIRECTORY, root));
  const visited = new Set<string>();
  const violations: NeutralBoundaryViolation[] = [];

  while (pending.length > 0) {
    const file = pending.pop();
    if (file === undefined || visited.has(file)) continue;
    if (!existsSync(file)) {
      violations.push({
        file: relativeModuleName(file),
        line: 1,
        reason: "neutral comparison root or dependency does not exist",
      });
      continue;
    }
    visited.add(file);
    const analyzed = analyzeSource(file, readFileSync(file, "utf8"));
    violations.push(...analyzed.violations);
    for (const specifier of analyzed.staticSpecifiers) {
      if (!specifier.startsWith(".")) {
        if (!ALLOWED_EXTERNAL_IMPORTS.has(specifier)) {
          violations.push({
            file: relativeModuleName(file),
            line: 1,
            reason: `external import ${JSON.stringify(specifier)} is not allowed`,
          });
        }
        continue;
      }
      const dependency = resolveRelativeModule(file, specifier);
      if (dependency === null) {
        violations.push({
          file: relativeModuleName(file),
          line: 1,
          reason: `relative import ${JSON.stringify(specifier)} cannot be resolved`,
        });
        continue;
      }
      const relativeDependency = relative(COMPARE_DIRECTORY, dependency);
      if (relativeDependency.startsWith("..") || path.isAbsolute(relativeDependency)) {
        violations.push({
          file: relativeModuleName(file),
          line: 1,
          reason: `relative import ${JSON.stringify(specifier)} escapes the neutral compare domain`,
        });
        continue;
      }
      pending.push(dependency);
    }
  }

  return {
    modules: [...visited].map(relativeModuleName).toSorted(),
    violations,
  };
};

describe("neutral comparison dependency boundary", () => {
  test("the graph derived from public roots stays pure and deterministic", () => {
    const graph = neutralModuleGraph();

    expect(graph.modules).toContain("content.ts");
    expect(graph.modules).toContain("content-types.ts");
    expect(graph.violations).toEqual([]);
  });

  test("the analyzer rejects deferred dependencies and ambient runtime state", () => {
    const synthetic = analyzeSource(
      resolve(COMPARE_DIRECTORY, "synthetic.ts"),
      [
        'const lazy = import("./docx-adapter");',
        'const loaded = require("./browser-adapter");',
        "const started = Date.now();",
        "const elapsed = performance.now();",
        "const nonce = Math.random();",
        "const body = document.body;",
        'const ordered = "a".localeCompare("b");',
      ].join("\n"),
    );

    expect(synthetic.violations.map(({ reason }) => reason).toSorted()).toEqual(
      [
        "CommonJS require is not allowed",
        "Math.random is not allowed",
        "ambient-locale operations are not allowed",
        "dynamic import is not allowed",
        'runtime global "Date" is not allowed',
        'runtime global "document" is not allowed',
        'runtime global "performance" is not allowed',
      ].toSorted(),
    );
  });

  test("property names, type names, and string contents do not false-fire", () => {
    const synthetic = analyzeSource(
      resolve(COMPARE_DIRECTORY, "synthetic.ts"),
      [
        "type Document = { readonly performance: string };",
        'const record = { document: "Math.random()" };',
        "const read = (value: Document): string => value.performance;",
        "void record;",
        "void read;",
      ].join("\n"),
    );

    expect(synthetic.violations).toEqual([]);
  });
});
