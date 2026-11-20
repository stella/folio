/**
 * A canonical session names a text box's paragraphs through the story drawing
 * the box. `storyParagraphs` lists a story's own paragraphs and none in its
 * text boxes, so canonical session code that looks a paragraph up by it misses
 * every box paragraph. It reads them through `paragraphsThroughTextBoxes`.
 */
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId: "throughTextBoxes";
  }) => void;
};

const BANNED = "storyParagraphs";

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";
const nameOf = (value: unknown): string | undefined => {
  if (!isNode(value)) return undefined;
  if (value.type === "Identifier" && typeof value.name === "string")
    return value.name;
  if (value.type === "Literal" && typeof value.value === "string")
    return value.value;
  return undefined;
};

export default {
  meta: { name: "folio-story-paragraphs" },
  rules: {
    "through-text-boxes": {
      meta: {
        type: "problem",
        messages: {
          throughTextBoxes:
            "Canonical session code reads a story's paragraphs through paragraphsThroughTextBoxes; storyParagraphs misses the paragraphs of its text boxes.",
        },
      },
      create(context: RuleContext) {
        return {
          // A named import, renamed or not.
          ImportSpecifier: (node: AstNode) => {
            if (nameOf(node.imported) === BANNED)
              context.report({ node, messageId: "throughTextBoxes" });
          },
          // A re-export hands it on to the same code.
          ExportSpecifier: (node: AstNode) => {
            if (nameOf(node.local) === BANNED)
              context.report({ node, messageId: "throughTextBoxes" });
          },
          // A namespace or dynamic import read by name.
          MemberExpression: (node: AstNode) => {
            const property = node.property;
            if (
              (node.computed !== true && nameOf(property) === BANNED) ||
              (node.computed === true &&
                isNode(property) &&
                property.type === "Literal" &&
                property.value === BANNED)
            )
              context.report({ node, messageId: "throughTextBoxes" });
          },
          // A destructured namespace or dynamic import.
          Property: (node: AstNode) => {
            const parent = node.parent;
            if (
              isNode(parent) &&
              parent.type === "ObjectPattern" &&
              nameOf(node.key) === BANNED
            )
              context.report({ node, messageId: "throughTextBoxes" });
          },
        };
      },
    },
  },
};
