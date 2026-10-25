type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId:
      "directTextShape" | "positionedAdvanceOwner" | "linePmRangeOwner";
  }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" && value !== null && "type" in value;

const identifierName = (node: unknown): string | null =>
  isAstNode(node) &&
  node.type === "Identifier" &&
  typeof node["name"] === "string"
    ? node["name"]
    : null;

const literalName = (node: unknown): string | null =>
  isAstNode(node) &&
  node.type === "Literal" &&
  typeof node["value"] === "string"
    ? node["value"]
    : null;

const propertyName = (node: AstNode): string | null =>
  node["computed"] === true
    ? literalName(node.property)
    : identifierName(node.property);

const isForbiddenMember = (node: AstNode): boolean => {
  const property = propertyName(node);
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
    "positioned-advance-owner": {
      meta: {
        type: "problem",
        messages: {
          positionedAdvanceOwner:
            "Glyph and text-position geometry must consume measurePositionedRun so shaping cannot rescale automatic advances or omit boundary advances.",
          linePmRangeOwner:
            "Text-position geometry must import computeLinePmRange from linePmRange instead of duplicating PM-boundary arithmetic.",
        },
      },
      create(context: RuleContext) {
        return {
          FunctionDeclaration: (node: unknown) => {
            if (
              isAstNode(node) &&
              identifierName(node["id"]) === "computeLinePmRange"
            )
              context.report({ node, messageId: "linePmRangeOwner" });
          },
          VariableDeclarator: (node: unknown) => {
            if (
              isAstNode(node) &&
              identifierName(node["id"]) === "computeLinePmRange"
            )
              context.report({ node, messageId: "linePmRangeOwner" });
          },
          ImportDeclaration: (node: unknown) => {
            if (!isAstNode(node)) return;
            const source = literalName(node.source);
            if (source === null || !source.endsWith("/measureProvider")) return;
            const specifiers = node["specifiers"];
            if (!Array.isArray(specifiers)) return;
            for (const specifier of specifiers) {
              if (!isAstNode(specifier)) continue;
              const imported =
                identifierName(specifier["imported"]) ??
                literalName(specifier["imported"]);
              if (
                specifier.type === "ImportNamespaceSpecifier" ||
                imported === "measureRun" ||
                imported === "measureText" ||
                imported === "measureTextWidth"
              )
                context.report({
                  node: specifier,
                  messageId: "positionedAdvanceOwner",
                });
            }
          },
        };
      },
    },
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
