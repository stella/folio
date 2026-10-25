import type { Node, Reference, Rule } from "@oxlint/plugins";
import { panic } from "better-result";
import path from "node:path";
import { fileURLToPath } from "node:url";

import baseline from "./folio-test-dom.baseline.json" with { type: "json" };

const BASELINE: Readonly<Record<string, number>> = baseline;
const REPO_ROOT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const REGISTRATOR_PACKAGE = "@happy-dom/global-registrator";
const isRegistratorModule = (value: unknown): boolean =>
  typeof value === "string" &&
  (value === REGISTRATOR_PACKAGE ||
    value.startsWith(REGISTRATOR_PACKAGE + "/"));
const TYPE_CONTEXTS = new Set([
  "TSTypeQuery",
  "TSTypeReference",
  "TSTypeAnnotation",
  "TSImportType",
  "TSInterfaceHeritage",
  "TSClassImplements",
]);

const isValueReference = (reference: Reference): boolean => {
  if (!reference.isRead()) return false;
  let node: Node | null = reference.identifier;
  while (node !== null) {
    if (TYPE_CONTEXTS.has(node.type)) return false;
    node = node.parent;
  }
  return true;
};

const domLeaseOwner = {
  meta: {
    type: "problem",
    messages: {
      domLeaseOwner:
        "Borrow DOM globals through acquireHarnessDom from test/harnessDom and release the lease in finally. Raw provider references can transfer global DOM ownership; legacy references may only shrink.",
      staleBaseline:
        "Reduce this file's folio-test-dom baseline after removing raw provider references; unused allowances must not survive a migration.",
    },
  },
  create(context) {
    const filename = path
      .relative(REPO_ROOT, context.filename)
      .replaceAll("\\", "/");
    const allowed = BASELINE[filename] ?? 0;
    let seen = 0;
    const reportReference = (node: Node) => {
      seen++;
      if (seen > allowed) context.report({ node, messageId: "domLeaseOwner" });
    };
    return {
      ImportDeclaration(node) {
        if (
          !isRegistratorModule(node.source.value) ||
          node.importKind === "type"
        )
          return;
        for (const variable of context.sourceCode.getDeclaredVariables(node)) {
          if (
            variable.defs.some(
              (definition) =>
                definition.node.type === "ImportSpecifier" &&
                definition.node.importKind === "type",
            )
          )
            continue;
          for (const reference of variable.references) {
            if (isValueReference(reference))
              reportReference(reference.identifier);
          }
        }
      },
      ImportExpression(node) {
        if (
          node.source.type === "Literal" &&
          isRegistratorModule(node.source.value)
        )
          reportReference(node);
      },
      CallExpression(node) {
        if (node.callee.type !== "Identifier" || node.callee.name !== "require")
          return;
        const reference = context.sourceCode
          .getScope(node.callee)
          .references.find((candidate) => candidate.identifier === node.callee);
        if (reference === undefined)
          panic("Missing native scope reference for require");
        if (
          reference.resolved?.defs.some(
            (definition) => definition.type !== "ImplicitGlobalVariable",
          )
        )
          return;
        const first = node.arguments.at(0);
        if (first?.type === "Literal" && isRegistratorModule(first.value))
          reportReference(node);
      },
      "Program:exit"(node) {
        if (seen < allowed)
          context.report({ node, messageId: "staleBaseline" });
      },
    };
  },
} satisfies Rule;

export default {
  meta: { name: "folio-test-dom" },
  rules: { "dom-lease-owner": domLeaseOwner },
};
