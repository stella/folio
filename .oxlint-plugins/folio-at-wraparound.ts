// Forbid `.at()` reads whose index is computed by subtraction or negation.
//
// `.at()` counts a negative index from the end, so `list.at(index - 1)` at
// `index === 0` reads the list's last item instead of nothing, and
// `list.at(-count)` at `count === 0` reads its first. Both look like a
// neighbour read and fail only at the edge.
//
// Each package reads the item before an index through its
// `itemBefore(list, index)` helper, which returns `undefined` at the list's
// start; `scripts/codemod-at-wraparound.ts` rewrites `.at(index - 1)` reads to
// it. Other computed indexes check their sign first. A literal negative index
// (`list.at(-1)`) names an item from the end on purpose and stays allowed, as
// does a cyclic read under a disable comment that says so.
//
// Flagged:
//   list.at(index - 1)
//   list.at(end - start)
//   list.at(-count)
//
// Not flagged:
//   list.at(-1)
//   list.at(index + 1)
//   itemBefore(list, index)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "wrappingIndex" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const WRAPPERS = new Set(["ParenthesizedExpression", "TSNonNullExpression"]);

const unwrap = (node: unknown): unknown =>
  isAstNode(node) && WRAPPERS.has(node.type)
    ? unwrap(node["expression"])
    : node;

// The name a member access reads: `a.name` or `a["name"]`.
const memberName = (node: AstNode): string | undefined => {
  const property = node["property"];
  if (!isAstNode(property)) {
    return undefined;
  }
  if (node["computed"] !== true) {
    return property.type === "Identifier"
      ? String(property["name"])
      : undefined;
  }
  return property.type === "Literal" && typeof property["value"] === "string"
    ? property["value"]
    : undefined;
};

/** How a `.at()` call computes a possibly negative index, if it does. */
export type WrappingIndex = "subtraction" | "negation";

/**
 * Whether `node` is a `.at()` call whose index the rule flags, and how that
 * index is computed. `scripts/codemod-at-wraparound.ts` selects with it too.
 */
export const wrappingIndexOf = (node: unknown): WrappingIndex | undefined => {
  if (!isAstNode(node) || node.type !== "CallExpression") {
    return undefined;
  }
  const { callee, arguments: args } = node;
  if (!isAstNode(callee) || callee.type !== "MemberExpression") {
    return undefined;
  }
  if (memberName(callee) !== "at" || !Array.isArray(args)) {
    return undefined;
  }
  const index = unwrap(args.at(0));
  if (!isAstNode(index)) {
    return undefined;
  }
  if (index.type === "BinaryExpression" && index["operator"] === "-") {
    return "subtraction";
  }
  const negated = unwrap(index["argument"]);
  return index.type === "UnaryExpression" &&
    index["operator"] === "-" &&
    !(isAstNode(negated) && negated.type === "Literal")
    ? "negation"
    : undefined;
};

export default {
  meta: { name: "folio-at-wraparound" },
  rules: {
    "no-computed-negative-index": {
      meta: {
        type: "problem",
        messages: {
          wrappingIndex:
            "`.at()` counts a negative index from the end, so a computed " +
            "index that goes below zero reads an item from the other end of " +
            "the list. Read the item before an index with the package's " +
            "`itemBefore(list, index)`, which returns `undefined` at the " +
            "start; check any other computed index is not negative first.",
        },
      },
      create(context: RuleContext) {
        return {
          CallExpression: (node: unknown) => {
            if (!isAstNode(node) || wrappingIndexOf(node) === undefined) {
              return;
            }
            // The `at` name: a disable comment above a chained `.at()` line
            // covers it.
            const callee = node["callee"];
            const property = isAstNode(callee) ? callee["property"] : node;
            context.report({ node: property, messageId: "wrappingIndex" });
          },
        };
      },
    },
  },
};
