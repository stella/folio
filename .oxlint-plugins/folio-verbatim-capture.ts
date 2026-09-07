// Content the DOCX parser preserves verbatim reaches a rebuilt part as a
// string, and a rebuilt part is Transitional. `captureVerbatimXml` converts a
// Strict fragment on the way out; `elementToXml` does not, so a parser that
// calls it directly replays Strict markup under a Transitional root.
//
// `metadataPrivacy` is the one caller that must not convert: it rewrites
// `docProps/core.xml` in place, under the root that part already has.

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  filename: string;
  report: (descriptor: { node: unknown; messageId: "directElementToXml" }) => void;
};

const OWNERS = ["docx/verbatimCapture.ts", "docx/metadataPrivacy.ts"];

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const importsElementToXml = (node: AstNode): boolean => {
  const source = node.source;
  if (!isAstNode(source) || typeof source.value !== "string") {
    return false;
  }
  if (!source.value.endsWith("xmlParser")) {
    return false;
  }
  const specifiers = node.specifiers;
  if (!Array.isArray(specifiers)) {
    return false;
  }
  return specifiers.some((specifier) => {
    if (!isAstNode(specifier) || specifier.type !== "ImportSpecifier") {
      return false;
    }
    const imported = specifier.imported;
    return isAstNode(imported) && imported.name === "elementToXml";
  });
};

export default {
  meta: { name: "folio-verbatim-capture" },
  rules: {
    "no-direct-element-to-xml": {
      meta: {
        type: "problem",
        messages: {
          directElementToXml:
            "Capture replayed markup with `captureVerbatimXml` from `docx/verbatimCapture`. " +
            "`elementToXml` writes a fragment in the source's own conformance class, which " +
            "replays Strict namespaces and Strict value spellings under a Transitional root.",
        },
      },
      create(context: RuleContext) {
        const normalized = context.filename.replaceAll("\\", "/");
        if (OWNERS.some((owner) => normalized.endsWith(owner))) {
          return {};
        }
        return {
          ImportDeclaration: (node: unknown) => {
            if (isAstNode(node) && importsElementToXml(node)) {
              context.report({ node, messageId: "directElementToXml" });
            }
          },
        };
      },
    },
  },
};
