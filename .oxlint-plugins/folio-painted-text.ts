type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "directTextShape" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" && value !== null && "type" in value;

const identifierName = (node: unknown): string | null =>
  isAstNode(node) && node.type === "Identifier" && typeof node["name"] === "string"
    ? node["name"]
    : null;

const isForbiddenMember = (node: AstNode): boolean => {
  const property = identifierName(node.property);
  if (property === "firstChild") return true;

  const object = identifierName(node.object);
  return (
    (object === "Node" && property === "TEXT_NODE") ||
    (object === "NodeFilter" && property === "SHOW_TEXT")
  );
};

export default {
  meta: { name: "folio-painted-text" },
  rules: {
    "no-direct-text-shape": {
      meta: {
        type: "problem",
        messages: {
          directTextShape:
            "Painted text may contain arbitrary descendant wrappers. Use the textStreamDom helpers instead of assuming a direct text child or walking a private DOM shape.",
        },
      },
      create(context: RuleContext) {
        return {
          MemberExpression: (node: unknown) => {
            if (isAstNode(node) && isForbiddenMember(node)) {
              context.report({ node, messageId: "directTextShape" });
            }
          },
        };
      },
    },
  },
};
