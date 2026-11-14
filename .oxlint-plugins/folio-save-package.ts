// A save writes over the package it was opened from. The reader moves a main
// document the package names at another path to `word/document.xml`, and every
// writer works at that path, so a save entry reads the bytes it is handed
// through `readPackageForSave` from `docx/mainDocumentPart`, which moves the
// main document the same way. `ZipArchive.read` in a save module would hand
// the writer a package whose main document it does not find.

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId: "directPackageRead";
  }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const isNamed = (value: unknown, name: string): boolean =>
  isAstNode(value) && value.type === "Identifier" && value.name === name;

const readsZipArchive = (node: AstNode): boolean => {
  const callee = node.callee;
  return (
    isAstNode(callee) &&
    callee.type === "MemberExpression" &&
    isNamed(callee.object, "ZipArchive") &&
    isNamed(callee.property, "read")
  );
};

export default {
  meta: { name: "folio-save-package" },
  rules: {
    "read-through-owner": {
      meta: {
        type: "problem",
        messages: {
          directPackageRead:
            "Read the package a save writes over with `readPackageForSave` from " +
            "`docx/mainDocumentPart`. It puts the main document at the path every " +
            "writer uses; `ZipArchive.read` keeps it wherever the producer put it.",
        },
      },
      create(context: RuleContext) {
        return {
          CallExpression: (node: unknown) => {
            if (isAstNode(node) && readsZipArchive(node)) {
              context.report({ node, messageId: "directPackageRead" });
            }
          },
        };
      },
    },
  },
};
