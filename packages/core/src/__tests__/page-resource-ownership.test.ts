import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const CORE_SRC = path.resolve(import.meta.dir, "..");
const OWNER_PATH = path.resolve(
  CORE_SRC,
  "display-list/pageResourceOwnership.ts",
);
const FORBIDDEN_OWNER_IMPORT =
  /(?:^|\/)(?:dom|render-dom|layout-painter|test|tests)(?:\/|$)|\.test(?:\.|$)/u;

const importsForSource = (source: string): string[] => {
  const parsed = ts.createSourceFile(
    OWNER_PATH,
    source,
    ts.ScriptTarget.Latest,
    true,
  );
  return parsed.statements.flatMap((statement) =>
    ts.isImportDeclaration(statement) &&
    ts.isStringLiteral(statement.moduleSpecifier)
      ? [statement.moduleSpecifier.text]
      : [],
  );
};

const forbiddenOwnerImports = (source: string): string[] =>
  importsForSource(source).filter((specifier) =>
    FORBIDDEN_OWNER_IMPORT.test(specifier),
  );

const receiverOfMutation = (
  node: ts.Node,
  source: ts.SourceFile,
): string | undefined => {
  if (
    ts.isCallExpression(node) &&
    ts.isPropertyAccessExpression(node.expression) &&
    ["remove", "replaceWith"].includes(node.expression.name.text)
  ) {
    return node.expression.expression.getText(source);
  }
  if (
    ts.isBinaryExpression(node) &&
    node.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
    ts.isPropertyAccessExpression(node.left) &&
    node.left.name.text === "innerHTML"
  ) {
    return node.left.expression.getText(source);
  }
  return undefined;
};

const mutationsWithoutAdjacentRelease = (source: string): string[] => {
  const parsed = ts.createSourceFile(
    "page-lifecycle.ts",
    source,
    ts.ScriptTarget.Latest,
    true,
  );
  const violations: string[] = [];
  const visit = (node: ts.Node): void => {
    const receiver = receiverOfMutation(node, parsed);
    if (receiver !== undefined && ts.isExpressionStatement(node.parent)) {
      const block = node.parent.parent;
      const statements =
        ts.isBlock(block) || ts.isSourceFile(block) ? block.statements : [];
      if (statements.length > 0) {
        const index = statements.indexOf(node.parent);
        const previous = statements.at(index - 1);
        const expression = previous?.expression;
        const release =
          expression !== undefined && ts.isCallExpression(expression)
            ? expression
            : undefined;
        const releasedReceiver =
          release?.expression &&
          ts.isIdentifier(release.expression) &&
          release.expression.text === "releaseDisplayPageResources"
            ? release.arguments.at(0)?.getText(parsed)
            : undefined;
        if (releasedReceiver !== receiver) {
          violations.push(node.getText(parsed));
        }
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(parsed);
  return violations;
};

test("page resource ownership stays below renderers and DOM backends", () => {
  const source = readFileSync(OWNER_PATH, "utf8");
  expect(forbiddenOwnerImports(source)).toEqual([]);
});

test("the owner boundary rejects renderer, DOM, and test imports", () => {
  expect(
    forbiddenOwnerImports(
      [
        'import "../dom/renderDisplayListToDom";',
        'import type { RenderPage } from "../layout-painter/renderPage";',
        'import { fixture } from "./owner.test";',
      ].join("\n"),
    ),
  ).toEqual([
    "../dom/renderDisplayListToDom",
    "../layout-painter/renderPage",
    "./owner.test",
  ]);
});

test("page lifecycle callers import the shared owner directly", () => {
  for (const relativePath of [
    "display-list/dom/renderDisplayListToDom.ts",
    "layout-painter/index.ts",
    "layout-painter/renderPage.ts",
  ]) {
    const source = readFileSync(path.resolve(CORE_SRC, relativePath), "utf8");
    const parsed = ts.createSourceFile(
      relativePath,
      source,
      ts.ScriptTarget.Latest,
      true,
    );
    const importSpecifiers = parsed.statements.flatMap((statement) =>
      ts.isImportDeclaration(statement) &&
      ts.isStringLiteral(statement.moduleSpecifier)
        ? [statement.moduleSpecifier.text]
        : [],
    );
    expect(importSpecifiers).toContain("../display-list/pageResourceOwnership");
    expect(importSpecifiers).not.toContain(
      "../display-list/dom/renderDisplayListToDom",
    );
  }
});

test("page disposal releases resources from the exact receiver first", () => {
  for (const relativePath of [
    "layout-painter/index.ts",
    "layout-painter/renderPage.ts",
  ]) {
    const source = readFileSync(path.resolve(CORE_SRC, relativePath), "utf8");
    expect(mutationsWithoutAdjacentRelease(source)).toEqual([]);
  }
});

test("the page disposal guard rejects missing or mismatched owners", () => {
  expect(
    mutationsWithoutAdjacentRelease(
      [
        "const page = { remove() {} };",
        "releaseDisplayPageResources(other);",
        "page.remove();",
      ].join("\n"),
    ),
  ).toEqual(["page.remove()"]);
});
