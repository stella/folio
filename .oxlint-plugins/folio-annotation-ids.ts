// Forbid truthiness tests on an annotation id in folio source.
//
// Comments, tracked changes, bookmarks, footnotes and endnotes are numbered by
// `w:id`, and `0` is a real one: Word numbers a package's comments from it,
// and an editing session gives it to the first comment it creates in a
// document without any. A truthiness test reads that annotation as absent, so
// whatever it guards (a painted anchor, a sidebar card, a saved range)
// silently skips it. Absence is `undefined` or `null` and is tested as such.
//
// Flagged:
//   if (commentId) { ... }
//   const painted = mark.attrs.revisionId && paint(mark);
//   if (!activeCommentId) return;
//
// Not flagged:
//   if (commentId !== undefined) { ... }
//   if (commentIds.length > 0) { ... }   (a list, not an id)
//   if (styleId) { ... }                 (a name, not a number)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId: "annotationIdTruthiness";
  }) => void;
};

/**
 * `commentId`, `activeCommentId`, `attrs.revisionId`, `footnoteRefId`, never
 * `commentIds`: the numbered annotation families, whatever prefixes them.
 */
const ANNOTATION_ID_NAME =
  /(?:^|[a-z])(?:[cC]omment|[rR]evision|[bB]ookmark|[fF]ootnote(?:Ref)?|[eE]ndnote(?:Ref)?|[nN]ote)Id$/u;

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const namesAnnotationId = (value: unknown): boolean => {
  if (!isAstNode(value)) {
    return false;
  }
  switch (value.type) {
    case "Identifier":
      return (
        typeof value["name"] === "string" &&
        ANNOTATION_ID_NAME.test(value["name"])
      );
    case "MemberExpression": {
      const property = value["property"];
      return (
        value["computed"] !== true &&
        isAstNode(property) &&
        typeof property["name"] === "string" &&
        ANNOTATION_ID_NAME.test(property["name"])
      );
    }
    case "ChainExpression":
    case "TSNonNullExpression":
      return namesAnnotationId(value["expression"]);
    default:
      return false;
  }
};

/** The annotation ids an expression tests for truth when it is itself a test. */
const testedAnnotationIds = (value: unknown): AstNode[] => {
  if (!isAstNode(value)) {
    return [];
  }
  if (namesAnnotationId(value)) {
    return [value];
  }
  if (value.type === "UnaryExpression" && value["operator"] === "!") {
    return testedAnnotationIds(value["argument"]);
  }
  if (
    value.type === "LogicalExpression" &&
    (value["operator"] === "&&" || value["operator"] === "||")
  ) {
    return [
      ...testedAnnotationIds(value["left"]),
      ...testedAnnotationIds(value["right"]),
    ];
  }
  return [];
};

export default {
  meta: { name: "folio-annotation-ids" },
  rules: {
    "no-annotation-id-truthiness": {
      meta: {
        type: "problem",
        messages: {
          annotationIdTruthiness:
            "Compare an annotation id with `undefined` or `null`. `0` is a real `w:id`, " +
            "the first one a package or an editing session numbers, and a truthiness " +
            "test reads it as absent.",
        },
      },
      create(context: RuleContext) {
        // `if (!commentId)` is both a test and a negation; report it once.
        const reported = new Set<unknown>();
        const report = (node: AstNode) => {
          const range = node["range"];
          const start = Array.isArray(range) ? range[0] : node["start"];
          const key = typeof start === "number" ? start : node;
          if (reported.has(key)) return;
          reported.add(key);
          context.report({ node, messageId: "annotationIdTruthiness" });
        };
        const reportTest = (test: unknown) => {
          for (const node of testedAnnotationIds(test)) report(node);
        };
        const reportOperand = (operand: unknown) => {
          if (isAstNode(operand) && namesAnnotationId(operand)) report(operand);
        };
        return {
          IfStatement: (node: AstNode) => reportTest(node["test"]),
          ConditionalExpression: (node: AstNode) => reportTest(node["test"]),
          WhileStatement: (node: AstNode) => reportTest(node["test"]),
          DoWhileStatement: (node: AstNode) => reportTest(node["test"]),
          ForStatement: (node: AstNode) => reportTest(node["test"]),
          // Outside a test, `!id` and the left side of `&&` / `||` still read
          // the id as a boolean; a test above already reported its own.
          UnaryExpression: (node: AstNode) => {
            if (node["operator"] === "!") reportOperand(node["argument"]);
          },
          LogicalExpression: (node: AstNode) => {
            if (node["operator"] === "&&" || node["operator"] === "||") {
              reportOperand(node["left"]);
            }
          },
        };
      },
    },
  },
};
