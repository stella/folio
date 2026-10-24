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

const moduleSpecifiersForSource = (source: string): string[] => {
  const parsed = ts.createSourceFile(
    OWNER_PATH,
    source,
    ts.ScriptTarget.Latest,
    true,
  );
  const specifiers: string[] = [];
  const visit = (node: ts.Node): void => {
    if (
      (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) &&
      node.moduleSpecifier !== undefined &&
      ts.isStringLiteral(node.moduleSpecifier)
    ) {
      specifiers.push(node.moduleSpecifier.text);
    }
    const firstArgument =
      ts.isCallExpression(node) &&
      node.expression.kind === ts.SyntaxKind.ImportKeyword
        ? node.arguments.at(0)
        : undefined;
    if (firstArgument !== undefined && ts.isStringLiteral(firstArgument)) {
      specifiers.push(firstArgument.text);
    }
    ts.forEachChild(node, visit);
  };
  visit(parsed);
  return specifiers;
};

const forbiddenOwnerImports = (source: string): string[] =>
  moduleSpecifiersForSource(source).filter((specifier) =>
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
        'export * from "../render-dom/commentAnchorAttributes";',
        'void import("../layout-painter/renderPage");',
      ].join("\n"),
    ),
  ).toEqual([
    "../dom/renderDisplayListToDom",
    "../layout-painter/renderPage",
    "./owner.test",
    "../render-dom/commentAnchorAttributes",
    "../layout-painter/renderPage",
  ]);
});

test("page lifecycle callers import the shared owner directly", () => {
  const callerEntries = [
    {
      path: "display-list/dom/renderDisplayListToDom.ts",
      ownerImport: "../pageResourceOwnership",
    },
    {
      path: "layout-painter/index.ts",
      ownerImport: "../display-list/pageResourceOwnership",
    },
    {
      path: "layout-painter/renderPage.ts",
      ownerImport: "../display-list/pageResourceOwnership",
    },
  ];
  for (const callers of callerEntries) {
    const source = readFileSync(path.resolve(CORE_SRC, callers.path), "utf8");
    const parsed = ts.createSourceFile(
      callers.path,
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
    expect(importSpecifiers).toContain(callers.ownerImport);
    expect(importSpecifiers).not.toContain(
      "../display-list/dom/renderDisplayListToDom",
    );
  }
});

test("framework adapters release page resources through the public owner", () => {
  const adapterCases = [
    ["../../react/src/paged-editor/PagedEditor.tsx", "pagesContainer"],
    ["../../vue/src/composables/useDocxEditor.ts", "pagesContainer.value"],
  ] as const;
  for (const [relativePath, receiver] of adapterCases) {
    const sourcePath = path.resolve(CORE_SRC, relativePath);
    const source = readFileSync(sourcePath, "utf8");
    const parsed = ts.createSourceFile(
      sourcePath,
      source,
      ts.ScriptTarget.Latest,
      true,
    );
    const hasPublicOwnerImport = parsed.statements.some(
      (statement) =>
        ts.isImportDeclaration(statement) &&
        ts.isStringLiteral(statement.moduleSpecifier) &&
        statement.moduleSpecifier.text === "@stll/folio-core/layout-painter" &&
        statement.importClause?.namedBindings !== undefined &&
        ts.isNamedImports(statement.importClause.namedBindings) &&
        statement.importClause.namedBindings.elements.some(
          (binding) =>
            (binding.propertyName ?? binding.name).text ===
            "releaseDisplayPageResources",
        ),
    );
    const hasMatchingReleaseCall = parsed.statements.some((statement) => {
      let found = false;
      const visit = (node: ts.Node): void => {
        if (
          ts.isCallExpression(node) &&
          ts.isIdentifier(node.expression) &&
          node.expression.text === "releaseDisplayPageResources" &&
          node.arguments.at(0)?.getText(parsed) === receiver
        ) {
          found = true;
        }
        ts.forEachChild(node, visit);
      };
      visit(statement);
      return found;
    });
    expect(hasPublicOwnerImport, relativePath).toBe(true);
    expect(hasMatchingReleaseCall, relativePath).toBe(true);
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
