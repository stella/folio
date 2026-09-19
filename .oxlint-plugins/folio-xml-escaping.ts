// Forbid a hand-rolled XML escaper anywhere folio writes a package part.
//
// `escapeXmlText` and `escapeXmlAttribute` in
// `packages/docx-core/src/serialize/xmlEscape.ts` own this. Six copies of the
// escape chain existed before that module, and no two agreed: some escaped
// three metacharacters, some five; none dropped the characters XML 1.0 §2.2
// forbids, and none wrote the character references §3.3.3 and §2.11 require
// for a tab, a newline or a carriage return. A writer that escapes by hand
// disagrees with the owner about one of those, and the disagreement is a
// package Word refuses to open or silently repairs.
//
// The shape this recognises is the one every copy had: an XML entity
// replacement written as a string literal, either as the replacement argument
// of `.replace` / `.replaceAll` or as a value in an escape table.
//
// Flagged:
//   value.replace(/&/g, "&amp;").replace(/</g, "&lt;")
//   value.replaceAll("<", "&lt;")
//   const ESCAPES = { "&": "&amp;", "<": "&lt;" };
//
// Safe:
//   escapeXmlText(value)
//   escapeXmlAttribute(value)
//   xml.includes("&amp;")               (reading, not escaping)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  filename: string;
  report: (descriptor: { node: unknown; messageId: "handRolledXmlEscape" }) => void;
};

const OWNER = "packages/docx-core/src/serialize/xmlEscape.ts";

/** The five predefined XML entities, spelled as a replacement. */
const XML_ENTITY_REPLACEMENTS = new Set(["&amp;", "&lt;", "&gt;", "&quot;", "&apos;"]);

const REPLACE_METHODS = new Set(["replace", "replaceAll"]);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const xmlEntityLiteral = (node: unknown): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  if (node.type === "Literal") {
    return typeof node.value === "string" && XML_ENTITY_REPLACEMENTS.has(node.value);
  }
  // An untagged template with no substitutions is still a string literal.
  if (node.type !== "TemplateLiteral" || (node.expressions as unknown[])?.length !== 0) {
    return false;
  }
  const quasis = node.quasis;
  if (!Array.isArray(quasis) || quasis.length !== 1) {
    return false;
  }
  const cooked = (quasis[0] as { value?: { cooked?: unknown } }).value?.cooked;
  return typeof cooked === "string" && XML_ENTITY_REPLACEMENTS.has(cooked);
};

const isReplaceCall = (node: AstNode): boolean => {
  const callee = node.callee;
  if (!isAstNode(callee) || callee.type !== "MemberExpression") {
    return false;
  }
  const property = callee.property;
  return (
    isAstNode(property) &&
    property.type === "Identifier" &&
    typeof property["name"] === "string" &&
    REPLACE_METHODS.has(property["name"])
  );
};

export default {
  meta: { name: "folio-xml-escaping" },
  rules: {
    "no-hand-rolled-xml-escape": {
      meta: {
        type: "problem",
        messages: {
          handRolledXmlEscape:
            "Escape XML through `escapeXmlText` or `escapeXmlAttribute` from " +
            "`@stll/docx-core`. A second escaper always disagrees with the owner about " +
            "some character — the XML 1.0 §2.2 characters that must be dropped rather " +
            "than escaped, or the tab, newline and carriage return that need a " +
            "character reference to survive §3.3.3 and §2.11 — and the disagreement is " +
            "a package Word refuses to open.",
        },
      },
      create(context: RuleContext) {
        if (context.filename.replaceAll("\\", "/").endsWith(OWNER)) {
          return {};
        }
        return {
          CallExpression: (node: unknown) => {
            if (!isAstNode(node) || !isReplaceCall(node)) {
              return;
            }
            const args = Array.isArray(node.arguments) ? node.arguments : [];
            if (args.some(xmlEntityLiteral)) {
              context.report({ node, messageId: "handRolledXmlEscape" });
            }
          },
          Property: (node: unknown) => {
            if (isAstNode(node) && xmlEntityLiteral(node.value)) {
              context.report({ node, messageId: "handRolledXmlEscape" });
            }
          },
        };
      },
    },
  },
};
