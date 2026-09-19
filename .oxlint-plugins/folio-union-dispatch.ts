// A chain of `else if` over a discriminated union's tag is a silent-drift
// mechanism: a member added to the union later reaches whichever branch
// happens to be last (or no branch at all) and the build still passes. The
// container survival census records the result as a loss, months after the
// fact. A `switch` over the same tag with a `never` default turns that into a
// compile error at the moment the member is added.
//
// The rule has no type information, so it cannot tell a model union's tag from
// any other `.type` string. Enrollment is therefore the `files` glob of its
// `overrides` entry in `oxlint.config.ts`: the directories whose chains have
// been converted. That list may grow and may not shrink, which is what makes
// the remaining chains a fixed set.

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "elseIfChainOverTag" }) => void;
};

/** Tags a model union discriminates on. A `.name`, `.kind` of a CSS value and
 * the like are not covered, because the rule cannot tell them apart. */
const TAGS = new Set(["type"]);

/** Below this a chain is a pair of cases, not a dispatch table. */
const MIN_BRANCHES = 3;

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const textOf = (node: unknown): string | null => {
  if (!isAstNode(node)) {
    return null;
  }
  if (node.type === "Identifier") {
    return typeof node.name === "string" ? node.name : null;
  }
  if (node.type === "ThisExpression") {
    return "this";
  }
  if (node.type === "MemberExpression" || node.type === "ChainExpression") {
    const object = textOf(node.object ?? node.expression);
    const property = isAstNode(node.property) ? textOf(node.property) : null;
    return object && property ? `${object}.${property}` : null;
  }
  return null;
};

/** The subject of `x.y.type === "literal"`, or null when the test is anything else. */
const subjectOfTagTest = (test: unknown): string | null => {
  if (!isAstNode(test)) {
    return null;
  }
  if (test.type === "LogicalExpression" && (test.operator === "||" || test.operator === "&&")) {
    return subjectOfTagTest(test.left);
  }
  if (test.type !== "BinaryExpression" || test.operator !== "===") {
    return null;
  }
  const { left, right } = test;
  if (!isAstNode(left) || !isAstNode(right)) {
    return null;
  }
  if (right.type !== "Literal" || typeof right.value !== "string") {
    return null;
  }
  const member = left.type === "ChainExpression" ? left.expression : left;
  if (!isAstNode(member) || member.type !== "MemberExpression") {
    return null;
  }
  const tag = textOf(member.property);
  if (!tag || !TAGS.has(tag)) {
    return null;
  }
  return textOf(member.object);
};

/** A `const x: never = subject` (or an equivalent call) in the trailing else. */
const holdsNeverCheck = (node: unknown): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  if (
    node.type === "TSTypeAnnotation" &&
    isAstNode(node.typeAnnotation) &&
    node.typeAnnotation.type === "TSNeverKeyword"
  ) {
    return true;
  }
  return Object.values(node).some((value) =>
    Array.isArray(value) ? value.some(holdsNeverCheck) : holdsNeverCheck(value),
  );
};

export default {
  meta: { name: "folio-union-dispatch" },
  rules: {
    "exhaustive-model-union-dispatch": {
      meta: {
        type: "problem",
        messages: {
          elseIfChainOverTag:
            "Read a tagged union with a `switch` over its tag plus a `const _: never` default, " +
            "not a chain of `else if`. A chain lets a member added to the union later fall into " +
            "the last branch, or off the end, without failing the build.",
        },
      },
      create(context: RuleContext) {
        // A chain's tail is visited as an `IfStatement` of its own; report the
        // head once rather than every suffix of it.
        const tails = new WeakSet<object>();
        return {
          IfStatement: (node: unknown) => {
            if (!isAstNode(node) || tails.has(node)) {
              return;
            }
            const subject = subjectOfTagTest(node.test);
            if (!subject) {
              return;
            }
            let branches = 1;
            let alternate = node.alternate;
            while (isAstNode(alternate) && alternate.type === "IfStatement") {
              if (subjectOfTagTest(alternate.test) !== subject) {
                return;
              }
              tails.add(alternate);
              branches += 1;
              alternate = alternate.alternate;
            }
            if (branches < MIN_BRANCHES) {
              return;
            }
            if (alternate && holdsNeverCheck(alternate)) {
              return;
            }
            context.report({ node, messageId: "elseIfChainOverTag" });
          },
        };
      },
    },
  },
};
