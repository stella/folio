// Forbid passing an array to `Math.max` or `Math.min` as call arguments.
//
// `Math.max(...values)` turns every element into an argument, and engines
// throw `RangeError` once the array outgrows their argument limit (about a
// hundred thousand elements in some engines, fewer on small stacks). Document
// content sets these sizes, so the call fails on large input only. `apply`
// and `Reflect.apply` hand over the same argument list.
//
// Each package folds instead through its `maxOf` / `minOf` helper
// (`maxOf(values, floor)` is `Math.max(floor, ...values)`).
// `scripts/codemod-math-extremes.ts` rewrites spread calls to them.
//
// Flagged:
//   Math.max(...values)
//   Math.min(0, ...values)
//   Math.max.apply(null, values)
//   Reflect.apply(Math.min, undefined, values)
//
// Not flagged:
//   Math.max(a, b)
//   maxOf(values, 0)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "spreadExtreme" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const EXTREMES = new Set(["max", "min"]);

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

// `Math`, `globalThis.Math`, `window.Math`.
const isMath = (node: unknown): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  if (node.type === "Identifier") {
    return node["name"] === "Math";
  }
  return node.type === "MemberExpression" && memberName(node) === "Math";
};

// `Math.max`, `Math.min`, `Math["max"]`.
const isMathExtreme = (node: unknown): boolean =>
  isAstNode(node) &&
  node.type === "MemberExpression" &&
  isMath(node["object"]) &&
  EXTREMES.has(memberName(node) ?? "");

const hasSpread = (args: unknown): boolean =>
  Array.isArray(args) &&
  args.some((arg) => isAstNode(arg) && arg.type === "SpreadElement");

const spreadsIntoExtreme = (call: AstNode): boolean => {
  const { callee, arguments: args } = call;
  if (!isAstNode(callee) || callee.type !== "MemberExpression") {
    return false;
  }
  if (isMathExtreme(callee)) {
    return hasSpread(args);
  }
  const method = memberName(callee);
  if (isMathExtreme(callee["object"])) {
    return method === "apply" || (method === "call" && hasSpread(args));
  }
  return (
    method === "apply" &&
    isAstNode(callee["object"]) &&
    callee["object"].type === "Identifier" &&
    callee["object"]["name"] === "Reflect" &&
    Array.isArray(args) &&
    isMathExtreme(args.at(0))
  );
};

export default {
  meta: { name: "folio-math-extremes" },
  rules: {
    "no-spread-arguments": {
      meta: {
        type: "problem",
        messages: {
          spreadExtreme:
            "Spreading an array into `Math.max` or `Math.min` passes each " +
            "element as an argument, and engines throw `RangeError` past " +
            "their argument limit. Use the package's `maxOf(values, floor)` " +
            "or `minOf(values, ceiling)`, which fold over the values.",
        },
      },
      create(context: RuleContext) {
        return {
          CallExpression: (node: unknown) => {
            if (isAstNode(node) && spreadsIntoExtreme(node)) {
              context.report({ node, messageId: "spreadExtreme" });
            }
          },
        };
      },
    },
  },
};
