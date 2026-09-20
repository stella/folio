// Resolve an identity-bearing attribute by namespace URI, never by prefix.
//
// `getAttribute(element, prefix, name)` tries `prefix:name`, then bare `name`,
// then any prefix that ends in `:name`. The last step is deliberate: a producer
// may bind WordprocessingML to `ns0:`, and without it those documents lose
// their attributes on round-trip.
//
// For a value that only describes its own element the fallback costs nothing.
// For a value that names something else it is a way to join the wrong two
// things: a `paraId` under a foreign prefix is returned as the paragraph
// identity the comment part points at, an unrelated `:id` as the relationship
// the image is stored under. `getAttributeByNamespaceUri` resolves the prefix
// against the element's namespace scope and cannot make that mistake.
//
// Which attributes carry an identity is derived, not listed: see
// `scripts/generate-identity-attributes.ts`.
//
// Flagged:
//   getAttribute(node, "w14", "paraId")
//   getAttribute(blip, "r", "embed")
//   getAttributeAnyPrefix(element, "paraId")
//
// Safe:
//   getAttributeByNamespaceUri(node, PARA_ID_NAMESPACE_URIS, "paraId")
//   getAttribute(child, null, "Id")        // unprefixed: no fallback runs
//   getAttribute(cols, "w", "space")       // column spacing, not xml:space
//   getAttribute(element, "w", "val")      // describes its own element

import { IDENTITY_ATTRIBUTES } from "../specifications/generated/identity-attributes.gen";

type AstNode = Record<string, unknown> & { type: string };

type IdentityReadContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId: "prefixResolvedIdentityRead";
    data: { attribute: string; reader: string; reason: string };
  }) => void;
};

/** Keys that never lead to child nodes (or lead back up the tree). */
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);

/**
 * Where each prefix-resolving reader takes its arguments.
 *
 * `namespaceArgument` is `null` for a reader that takes no prefix at all and
 * resolves by local name on its own, so every prefix is in scope for it.
 */
const PREFIX_READERS: Readonly<
  Record<string, { nameArgument: number; namespaceArgument: number | null }>
> = {
  getAttribute: { nameArgument: 2, namespaceArgument: 1 },
  getAttributeAnyPrefix: { nameArgument: 1, namespaceArgument: null },
  parseNumericAttribute: { nameArgument: 2, namespaceArgument: 1 },
  parseOnOffAttribute: { nameArgument: 2, namespaceArgument: 1 },
};

/**
 * `xmlParser` owns the resolution itself: it declares both the prefix readers
 * and `getAttributeByNamespaceUri`, and its own calls are how the fallback is
 * implemented.
 */
const OWNER_MODULES = ["packages/core/src/docx/xmlParser.ts"];

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const forEachChild = (node: AstNode, visit: (child: AstNode) => void): void => {
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) {
      continue;
    }
    if (Array.isArray(value)) {
      for (const item of value) {
        if (isAstNode(item)) {
          visit(item);
        }
      }
      continue;
    }
    if (isAstNode(value)) {
      visit(value);
    }
  }
};

const normalized = (filename: string): string => filename.replaceAll("\\", "/");

const isOwningModule = (filename: string): boolean => {
  const path = normalized(filename);
  return OWNER_MODULES.some((module) => path.endsWith(module));
};

/** The string a literal argument holds, or `null` for anything else. */
const stringArgument = (node: unknown): string | null => {
  if (!isAstNode(node) || node.type !== "Literal") {
    return null;
  }
  const value = node["value"];
  return typeof value === "string" ? value : null;
};

/**
 * The reader a call names, or `null`.
 *
 * A bare identifier only: `element.getAttribute("id")` is the DOM's own method,
 * which takes a qualified name and resolves nothing.
 */
const readerName = (node: AstNode): string | null => {
  const callee = node["callee"];
  if (!isAstNode(callee) || callee.type !== "Identifier") {
    return null;
  }
  const { name } = callee;
  return typeof name === "string" && name in PREFIX_READERS ? name : null;
};

const checkCall = (node: AstNode, context: IdentityReadContext): void => {
  const reader = readerName(node);
  if (reader === null) {
    return;
  }
  const signature = PREFIX_READERS[reader];
  if (signature === undefined) {
    return;
  }
  const args = node["arguments"];
  if (!Array.isArray(args)) {
    return;
  }

  const attribute = stringArgument(args[signature.nameArgument]);
  if (attribute === null) {
    return;
  }
  const identity = IDENTITY_ATTRIBUTES.get(attribute);
  if (identity === undefined) {
    return;
  }

  if (signature.namespaceArgument !== null) {
    const prefix = stringArgument(args[signature.namespaceArgument]);
    // A non-literal prefix is unreadable here, and `null` (or `""`) asks for an
    // unprefixed attribute, which never reaches the any-prefix fallback.
    if (prefix === null || prefix === "") {
      return;
    }
    if (identity.prefixes !== null && !identity.prefixes.includes(prefix)) {
      return;
    }
  }

  context.report({
    node,
    messageId: "prefixResolvedIdentityRead",
    data: { attribute, reader, reason: identity.reason },
  });
};

export default {
  meta: { name: "folio-identity-attributes" },
  rules: {
    "no-prefix-resolved-identity-read": {
      meta: {
        type: "problem",
        messages: {
          prefixResolvedIdentityRead:
            "`{{attribute}}` carries an identity ({{reason}}), and `{{reader}}` falls back to any " +
            "prefix ending in that local name, so a foreign attribute is read as the real one and " +
            "joins the wrong two things. Resolve it with `getAttributeByNamespaceUri(element, " +
            '<namespace URIs>, "{{attribute}}")`. The set is derived in ' +
            "`scripts/generate-identity-attributes.ts`.",
        },
      },
      create(context: IdentityReadContext) {
        return {
          Program: (node: unknown) => {
            if (!isAstNode(node) || isOwningModule(context.filename)) {
              return;
            }
            const visit = (child: AstNode): void => {
              if (child.type === "CallExpression") {
                checkCall(child, context);
              }
              forEachChild(child, visit);
            };
            visit(node);
          },
        };
      },
    },
  },
};
