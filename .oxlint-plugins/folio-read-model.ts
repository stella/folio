// Forbid casting a value to the model `Document` outside `readModel`.
//
// A host holds the editor's document model as `ReadonlyDocument`: in a
// canonical session it is the session's own document, shared with the
// session's caches and the host's undo history, so a write to it corrupts
// every holder at once. Readers that still take a plain `Document` (the style
// resolver, the numbering map, the hidden editor state) are given the model
// through `readModel` in `packages/core/src/types/documentModel.ts`, the one
// place that drops `readonly`, whose readers a deep-freeze test runs over
// frozen documents. Any other cast to `Document` silently hands a writer the
// model: the compiler can no longer tell a reader from a writer, and nothing
// copies first.
//
// Flagged (where `Document` is the model type imported from folio-core or
// docx-core):
//   model as Document
//   model as unknown as Document
//   <Document>model
//   value as Document | null
//
// Not flagged:
//   readModel(model)                         (the sanctioned reader cast)
//   cloneDocumentWithParagraphPropertySources(model)   (a writer's copy)
//   element as unknown as Document           (the DOM `Document`, not imported)
//   value satisfies Document                 (checks, does not cast)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId: "modelDocumentCast";
  }) => void;
};

const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

// Package specifiers that export the model `Document`.
const MODEL_PACKAGE_SPECIFIERS = new Set([
  "@stll/docx-core",
  "@stll/docx-core/model",
  "@stll/folio-core",
  "@stll/folio-core/types",
  "@stll/folio-core/types/document",
]);

// Relative specifiers are resolved against the importing file and compared
// with the modules under `packages/core/src/types` that export it.
const MODEL_MODULE_PATHS = ["", "/index", "/document"].map(
  (name) => `/packages/core/src/types${name}`,
);

const normalizePath = (segments: readonly string[]): string[] => {
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") {
      continue;
    }
    if (segment === "..") {
      out.pop();
      continue;
    }
    out.push(segment);
  }
  return out;
};

const isModelImportSource = (filename: string, specifier: string): boolean => {
  if (MODEL_PACKAGE_SPECIFIERS.has(specifier)) {
    return true;
  }
  if (!specifier.startsWith(".")) {
    return false;
  }
  const dir = normalizePath(filename.replaceAll("\\", "/").split("/")).slice(
    0,
    -1,
  );
  const resolved = `/${normalizePath([...dir, ...specifier.split("/")]).join("/")}`;
  const stripped = resolved.endsWith(".ts") ? resolved.slice(0, -3) : resolved;
  return MODEL_MODULE_PATHS.some((modulePath) => stripped.endsWith(modulePath));
};

const forEachChild = (node: AstNode, visit: (child: AstNode) => void): void => {
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) {
      continue;
    }
    if (Array.isArray(value)) {
      for (const item of value) {
        if (isAstNode(item)) {
          visit(item);
        }
      }
      continue;
    }
    if (isAstNode(value)) {
      visit(value);
    }
  }
};

/** The local names the file binds to the model `Document`. */
const modelDocumentNames = (
  program: AstNode,
  filename: string,
): Set<string> => {
  const names = new Set<string>();
  const body = program["body"];
  if (!Array.isArray(body)) {
    return names;
  }
  for (const statement of body) {
    if (!isAstNode(statement) || statement.type !== "ImportDeclaration") {
      continue;
    }
    const source = statement["source"];
    const specifiers = statement["specifiers"];
    if (
      !isAstNode(source) ||
      typeof source["value"] !== "string" ||
      !isModelImportSource(filename, source["value"]) ||
      !Array.isArray(specifiers)
    ) {
      continue;
    }
    for (const specifier of specifiers) {
      if (!isAstNode(specifier) || specifier.type !== "ImportSpecifier") {
        continue;
      }
      const imported = specifier["imported"];
      const local = specifier["local"];
      const importedName = isAstNode(imported)
        ? (imported["name"] ?? imported["value"])
        : undefined;
      if (
        importedName === "Document" &&
        isAstNode(local) &&
        typeof local["name"] === "string"
      ) {
        names.add(local["name"]);
      }
    }
  }
  return names;
};

/** Whether `node` names the model `Document`, alone or as a union member. */
const namesModelDocument = (node: unknown, names: Set<string>): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  if (node.type === "TSUnionType") {
    const members = node["types"];
    return (
      Array.isArray(members) &&
      members.some((member) => namesModelDocument(member, names))
    );
  }
  if (node.type !== "TSTypeReference") {
    return false;
  }
  const typeName = node["typeName"];
  return (
    isAstNode(typeName) &&
    typeName.type === "Identifier" &&
    typeof typeName["name"] === "string" &&
    names.has(typeName["name"])
  );
};

export default {
  meta: { name: "folio-read-model" },
  rules: {
    "no-model-document-cast": {
      meta: {
        type: "problem",
        messages: {
          modelDocumentCast:
            "Do not cast to the model `Document`. A host's document model is " +
            "`ReadonlyDocument`: the session, its caches and the undo history " +
            "share it, so a write through a cast corrupts all of them. Pass " +
            "the model to a reader with `readModel` from " +
            "`@stll/folio-core/types/documentModel` (the one sanctioned cast, " +
            "whose readers a deep-freeze test covers), or copy it with " +
            "`cloneDocumentWithParagraphPropertySources` before writing.",
        },
      },
      create(context: RuleContext) {
        let names = new Set<string>();
        const check = (node: unknown) => {
          if (
            isAstNode(node) &&
            names.size > 0 &&
            namesModelDocument(node["typeAnnotation"], names)
          ) {
            context.report({ node, messageId: "modelDocumentCast" });
          }
        };
        return {
          Program: (node: unknown) => {
            names = isAstNode(node)
              ? modelDocumentNames(node, context.filename)
              : new Set();
          },
          TSAsExpression: check,
          TSTypeAssertion: check,
        };
      },
    },
  },
};
