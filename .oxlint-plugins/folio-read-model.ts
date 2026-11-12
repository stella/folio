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
//   model as Document & { marker?: string }
//   model as Model.Document                 (import type * as Model from ...)
//   model as Writable                       (type Writable = Document, in the file)
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

type ModelNames = {
  /** Local names bound to the model `Document` itself. */
  documents: Set<string>;
  /** Local names of namespace imports of a module that exports it. */
  namespaces: Set<string>;
};

/** The local names through which the file reaches the model `Document`. */
const modelDocumentNames = (program: AstNode, filename: string): ModelNames => {
  const names: ModelNames = { documents: new Set(), namespaces: new Set() };
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
      if (!isAstNode(specifier)) {
        continue;
      }
      const local = specifier["local"];
      if (!isAstNode(local) || typeof local["name"] !== "string") {
        continue;
      }
      if (specifier.type === "ImportNamespaceSpecifier") {
        names.namespaces.add(local["name"]);
        continue;
      }
      if (specifier.type !== "ImportSpecifier") {
        continue;
      }
      const imported = specifier["imported"];
      const importedName = isAstNode(imported)
        ? (imported["name"] ?? imported["value"])
        : undefined;
      if (importedName === "Document") {
        names.documents.add(local["name"]);
      }
    }
  }
  return names;
};

const identifierName = (node: unknown): string | null =>
  isAstNode(node) &&
  node.type === "Identifier" &&
  typeof node["name"] === "string"
    ? node["name"]
    : null;

/** Whether a type reference's name is `Document` or `Namespace.Document`. */
const referencesModelDocument = (
  typeName: unknown,
  names: ModelNames,
): boolean => {
  const direct = identifierName(typeName);
  if (direct !== null) {
    return names.documents.has(direct);
  }
  if (!isAstNode(typeName) || typeName.type !== "TSQualifiedName") {
    return false;
  }
  const namespace = identifierName(typeName["left"]);
  return (
    namespace !== null &&
    names.namespaces.has(namespace) &&
    identifierName(typeName["right"]) === "Document"
  );
};

/**
 * Whether `node` names the model `Document`, alone or as a member of a union
 * or an intersection: each still hands out writable document fields.
 */
const namesModelDocument = (node: unknown, names: ModelNames): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  if (node.type === "TSParenthesizedType") {
    return namesModelDocument(node["typeAnnotation"], names);
  }
  if (node.type === "TSUnionType" || node.type === "TSIntersectionType") {
    const members = node["types"];
    return (
      Array.isArray(members) &&
      members.some((member) => namesModelDocument(member, names))
    );
  }
  return (
    node.type === "TSTypeReference" &&
    referencesModelDocument(node["typeName"], names)
  );
};

/** Every type alias declared anywhere in the file. */
const typeAliases = (node: unknown, found: AstNode[] = []): AstNode[] => {
  if (Array.isArray(node)) {
    for (const child of node) typeAliases(child, found);
    return found;
  }
  if (!isAstNode(node)) return found;
  if (node.type === "TSTypeAliasDeclaration") found.push(node);
  for (const [key, child] of Object.entries(node)) {
    if (key !== "parent" && typeof child === "object" && child !== null)
      typeAliases(child, found);
  }
  return found;
};

/**
 * Add each alias the file declares for the model `Document` (alone or in a
 * union or intersection), and aliases of those, to the names that reach it.
 */
const addModelDocumentAliases = (program: AstNode, names: ModelNames): void => {
  const aliases = typeAliases(program["body"]);
  let added = true;
  while (added) {
    added = false;
    for (const alias of aliases) {
      const name = identifierName(alias["id"]);
      if (
        name !== null &&
        !names.documents.has(name) &&
        namesModelDocument(alias["typeAnnotation"], names)
      ) {
        names.documents.add(name);
        added = true;
      }
    }
  }
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
        let names: ModelNames = {
          documents: new Set(),
          namespaces: new Set(),
        };
        const check = (node: unknown) => {
          if (
            isAstNode(node) &&
            (names.documents.size > 0 || names.namespaces.size > 0) &&
            namesModelDocument(node["typeAnnotation"], names)
          ) {
            context.report({ node, messageId: "modelDocumentCast" });
          }
        };
        return {
          Program: (node: unknown) => {
            if (isAstNode(node)) {
              names = modelDocumentNames(node, context.filename);
              addModelDocumentAliases(node, names);
            }
          },
          TSAsExpression: check,
          TSTypeAssertion: check,
        };
      },
    },
  },
};
