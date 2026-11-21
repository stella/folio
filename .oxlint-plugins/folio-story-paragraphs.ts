/**
 * A canonical session names a text box's paragraphs through the story drawing
 * the box. `storyParagraphs` and `locateParagraphInBlockList` read a story's
 * own paragraphs and none in its text boxes, so canonical session code that
 * looks a paragraph up by them misses every box paragraph. It reads them
 * through `paragraphsThroughTextBoxes` and `locateParagraphThroughTextBoxes`.
 */
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId: "throughTextBoxes";
  }) => void;
};

const BANNED: ReadonlySet<string> = new Set([
  "storyParagraphs",
  "locateParagraphInBlockList",
]);

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
            "Canonical session code reads a story's paragraphs through paragraphsThroughTextBoxes or locateParagraphThroughTextBoxes; storyParagraphs and locateParagraphInBlockList miss the paragraphs of its text boxes.",
        },
      },
      create(context: RuleContext) {
        return {
          // A named import, renamed or not.
          ImportSpecifier: (node: AstNode) => {
            if (BANNED.has(nameOf(node.imported) ?? ""))
              context.report({ node, messageId: "throughTextBoxes" });
          },
          // A re-export hands it on to the same code.
          ExportSpecifier: (node: AstNode) => {
            if (BANNED.has(nameOf(node.local) ?? ""))
              context.report({ node, messageId: "throughTextBoxes" });
          },
          // A namespace or dynamic import read by name.
          MemberExpression: (node: AstNode) => {
            const property = node.property;
            if (
              (node.computed !== true && BANNED.has(nameOf(property) ?? "")) ||
              (node.computed === true &&
                isNode(property) &&
                property.type === "Literal" &&
                typeof property.value === "string" &&
                BANNED.has(property.value))
            )
              context.report({ node, messageId: "throughTextBoxes" });
          },
          // A destructured namespace or dynamic import.
          Property: (node: AstNode) => {
            const parent = node.parent;
            if (
              isNode(parent) &&
              parent.type === "ObjectPattern" &&
              BANNED.has(nameOf(node.key) ?? "")
            )
              context.report({ node, messageId: "throughTextBoxes" });
          },
        };
      },
    },
  },
};
