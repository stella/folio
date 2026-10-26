import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  type Node,
  nodeText,
  parseSourceModule,
  walkAst,
} from "../../../../scripts/lib/oxc-ast";

const CORE_SRC = path.resolve(import.meta.dir, "..");
const OWNER_PATH = path.resolve(
  CORE_SRC,
  "display-list/pageResourceOwnership.ts",
);
const FORBIDDEN_OWNER_IMPORT =
  /(?:^|\/)(?:dom|render-dom|layout-painter|test|tests)(?:\/|$)|\.test(?:\.|$)/u;

const moduleSpecifiersForSource = (source: string): string[] => {
  const { program } = parseSourceModule(OWNER_PATH, source);
  const specifiers: string[] = [];
  walkAst(program, (node) => {
    if (
      (node.type === "ImportDeclaration" ||
        node.type === "ExportNamedDeclaration" ||
        node.type === "ExportAllDeclaration") &&
      node.source !== null
    ) {
      specifiers.push(node.source.value);
    }
    const firstArgument =
      node.type === "ImportExpression" ? node.source : undefined;
    if (
      firstArgument?.type === "Literal" &&
      typeof firstArgument.value === "string"
    ) {
      specifiers.push(firstArgument.value);
    }
  });
  return specifiers;
};

const forbiddenOwnerImports = (source: string): string[] =>
  moduleSpecifiersForSource(source).filter((specifier) =>
    FORBIDDEN_OWNER_IMPORT.test(specifier),
  );

const receiverOfMutation = (node: Node, text: string): string | undefined => {
  if (
    node.type === "CallExpression" &&
    node.callee.type === "MemberExpression" &&
    !node.callee.computed &&
    node.callee.property.type === "Identifier" &&
    ["remove", "replaceWith"].includes(node.callee.property.name)
  ) {
    return nodeText(text, node.callee.object);
  }
  if (
    node.type === "AssignmentExpression" &&
    node.operator === "=" &&
    node.left.type === "MemberExpression" &&
    !node.left.computed &&
    node.left.property.type === "Identifier" &&
    node.left.property.name === "innerHTML"
  ) {
    return nodeText(text, node.left.object);
  }
  return undefined;
};

// An optional chain wraps the call it ends in; the statement holds the chain.
const unwrapChain = (node: Node): Node =>
  node.type === "ChainExpression" ? node.expression : node;

const statementsOf = (node: Node | undefined): readonly Node[] =>
  node?.type === "Program" ||
  node?.type === "BlockStatement" ||
  node?.type === "StaticBlock"
    ? node.body
    : [];

const mutationsWithoutAdjacentRelease = (source: string): string[] => {
  const { program, text } = parseSourceModule("page-lifecycle.ts", source);
  const violations: string[] = [];
  walkAst(program, (node, ancestors) => {
    const receiver = receiverOfMutation(node, text);
    const parentAt = ancestors.at(-1)?.type === "ChainExpression" ? -2 : -1;
    const statement = ancestors.at(parentAt);
    if (receiver !== undefined && statement?.type === "ExpressionStatement") {
      const statements = statementsOf(ancestors.at(parentAt - 1));
      if (statements.length > 0) {
        const index = statements.indexOf(statement);
        const previous = index > 0 ? statements.at(index - 1) : undefined;
        const expression =
          previous?.type === "ExpressionStatement"
            ? unwrapChain(previous.expression)
            : undefined;
        const release =
          expression?.type === "CallExpression" ? expression : undefined;
        const releasedArgument =
          release?.callee.type === "Identifier" &&
          release.callee.name === "releaseDisplayPageResources"
            ? release.arguments.at(0)
            : undefined;
        const releasedReceiver =
          releasedArgument === undefined
            ? undefined
            : nodeText(text, releasedArgument);
        if (releasedReceiver !== receiver) {
          violations.push(nodeText(text, node));
        }
      }
    }
  });
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
    const { program } = parseSourceModule(callers.path, source);
    const importSpecifiers = program.body.flatMap((statement) =>
      statement.type === "ImportDeclaration" ? [statement.source.value] : [],
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
    const { program, text } = parseSourceModule(sourcePath, source);
    const hasPublicOwnerImport = program.body.some(
      (statement) =>
        statement.type === "ImportDeclaration" &&
        statement.source.value === "@stll/folio-core/layout-painter" &&
        statement.specifiers.some(
          (binding) =>
            binding.type === "ImportSpecifier" &&
            (binding.imported.type === "Identifier"
              ? binding.imported.name
              : binding.imported.value) === "releaseDisplayPageResources",
        ),
    );
    const hasMatchingReleaseCall = program.body.some((statement) => {
      let found = false;
      walkAst(statement, (node) => {
        const argument =
          node.type === "CallExpression" &&
          node.callee.type === "Identifier" &&
          node.callee.name === "releaseDisplayPageResources"
            ? node.arguments.at(0)
            : undefined;
        if (argument !== undefined && nodeText(text, argument) === receiver) {
          found = true;
        }
      });
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

test("a release after the first statement does not satisfy disposal ordering", () => {
  expect(
    mutationsWithoutAdjacentRelease(
      "page.remove(); releaseDisplayPageResources(page);",
    ),
  ).toEqual(["page.remove()"]);
});
