// Enforce folio's model-type completeness (issue #845).
//
// A parser dropped a field on the way through the model because it attached
// the field through a local intersection instead of declaring it on the
// shared model type (`ListRendering & { levelStarts: number[] }`), and a
// consumer then read the field with an `in` check
// (`"levelStarts" in paragraph.listRendering`) instead of a typed property
// access. Both patterns compile against a type the model does not actually
// have, so the inverse conversion silently drops the field: the compiler
// never sees that it is missing.
//
// Flagged examples:
//   const widened = (r: ListRendering): ListRendering & { levelStarts: number[] } => ...;
//                                        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ widens the model type locally
//   paragraph.listRendering && "levelStarts" in paragraph.listRendering
//                               ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ reads a field the type does not declare
//
// Safe examples:
//   const hasParaId = (b: Paragraph): b is Paragraph & { paraId: string } => ...;  // type-guard narrowing
//   paragraph.listRendering?.levelStarts;                                          // direct typed read
//   type Attr = ParagraphNumberingOverride & { readonly [BRAND]: true };           // phantom brand, no data

type AstNode = Record<string, unknown> & { type: string };

type WideningContext = {
  filename: string;
  getSourceCode?: () => { text?: unknown };
  report: (descriptor: {
    node: unknown;
    messageId: "intersectionWidening";
    data: { name: string };
  }) => void;
  sourceCode?: { text?: unknown };
};

type InCheckContext = {
  filename: string;
  getSourceCode?: () => { text?: unknown };
  report: (descriptor: {
    node: unknown;
    messageId: "inCheckOnModel";
    data: { key: string; name: string };
  }) => void;
  sourceCode?: { text?: unknown };
};

// Keys that never lead to child nodes (or lead back up the tree).
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);

// Generic wrappers whose first type argument stands in for the model type it
// wraps: `NonNullable<Paragraph>`, `Readonly<Paragraph>`.
const TRANSPARENT_GENERICS = new Set(["NonNullable", "Readonly"]);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

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

// The shared document model lives in `@stll/docx-core` and is re-exported
// from these modules under `packages/core/src/types`. Relative specifiers
// resolve against the importing file, so a subsystem-local `./types` (layout
// measures, AI-edit results, editor stories) is not a model import.
const MODEL_MODULE_DIR = "packages/core/src/types";
const MODEL_MODULE_PATHS = new Set(
  ["", "/index", "/document", "/content", "/formatting", "/colors"].map(
    (name) => `/${MODEL_MODULE_DIR}${name}`,
  ),
);
const MODEL_PACKAGE_SPECIFIERS = new Set(["@stll/docx-core", "@stll/docx-core/model"]);

const normalizePath = (segments: string[]): string[] => {
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

const resolvesToModelModule = (filename: string, specifier: string): boolean => {
  const dir = normalizePath(filename.replaceAll("\\", "/").split("/")).slice(0, -1);
  const resolved = normalizePath([...dir, ...specifier.split("/")]).join("/");
  // oxlint hands over repo-relative filenames; anchor so both spellings match.
  const stripped = `/${resolved.endsWith(".ts") ? resolved.slice(0, -3) : resolved}`;
  return [...MODEL_MODULE_PATHS].some((modulePath) => stripped.endsWith(modulePath));
};

const isModelImportSource = (filename: string, value: string): boolean =>
  MODEL_PACKAGE_SPECIFIERS.has(value) ||
  (value.startsWith(".") && resolvesToModelModule(filename, value));

/** The file's model type names: the local names of every specifier imported
 * (as a type or as a value; `import type { X }` and `import { type X }` both
 * qualify) from a model-type import source. */
const collectModelTypeNames = (program: AstNode, filename: string): Set<string> => {
  const names = new Set<string>();
  const visit = (node: AstNode): void => {
    if (node.type === "ImportDeclaration") {
      const source = node["source"];
      const specifiers = node["specifiers"];
      if (
        isAstNode(source) &&
        typeof source["value"] === "string" &&
        isModelImportSource(filename, source["value"]) &&
        Array.isArray(specifiers)
      ) {
        for (const specifier of specifiers) {
          if (!isAstNode(specifier) || specifier.type !== "ImportSpecifier") {
            continue;
          }
          const local = specifier["local"];
          if (isAstNode(local) && typeof local["name"] === "string") {
            names.add(local["name"]);
          }
        }
      }
    }
    forEachChild(node, visit);
  };
  visit(program);
  return names;
};

// --- Rule 1: no-model-intersection-widening -------------------------------

/** The model type name a `TSTypeReference` names, directly or through
 * `NonNullable<...>` / `Readonly<...>`, else `null`. */
const modelReferenceName = (node: unknown, modelTypeNames: Set<string>): string | null => {
  if (!isAstNode(node)) {
    return null;
  }
  // `Paragraph | undefined`: any member naming a model type counts.
  if (node.type === "TSUnionType") {
    const members = node["types"];
    if (!Array.isArray(members)) {
      return null;
    }
    for (const member of members) {
      const name = modelReferenceName(member, modelTypeNames);
      if (name !== null) {
        return name;
      }
    }
    return null;
  }
  if (node.type !== "TSTypeReference") {
    return null;
  }
  const typeName = node["typeName"];
  if (
    !isAstNode(typeName) ||
    typeName.type !== "Identifier" ||
    typeof typeName["name"] !== "string"
  ) {
    return null;
  }
  if (modelTypeNames.has(typeName["name"])) {
    return typeName["name"];
  }
  if (!TRANSPARENT_GENERICS.has(typeName["name"])) {
    return null;
  }
  const typeArguments = node["typeArguments"];
  const params = isAstNode(typeArguments) ? typeArguments["params"] : null;
  const first = Array.isArray(params) ? params[0] : null;
  return modelReferenceName(first, modelTypeNames);
};

/** True when `node` is, through zero or more `TSTypeAnnotation` wrappers, the
 * `typeAnnotation` of a `TSTypePredicate` — a type-guard return type such as
 * `(b): b is Paragraph & { paraId: string }`, which narrows an existing
 * optional field rather than widening the model type. */
const isTypePredicateAnnotation = (node: AstNode): boolean => {
  let ancestor = node["parent"];
  while (isAstNode(ancestor) && ancestor.type === "TSTypeAnnotation") {
    ancestor = ancestor["parent"];
  }
  return isAstNode(ancestor) && ancestor.type === "TSTypePredicate";
};

/** The locally declared `unique symbol` names that can key phantom brands. */
const uniqueSymbolNames = (context: WideningContext): Set<string> => {
  const source = sourceTextForContext(context);
  return new Set(
    [
      ...source.matchAll(/\b(?:declare\s+)?const\s+(?<name>[$\w]+)\s*:\s*unique\s+symbol\b/gu),
    ].flatMap(({ groups }) => (groups?.["name"] === undefined ? [] : [groups["name"]])),
  );
};

/** True when every member is keyed by a declared `unique symbol`: a phantom
 * brand. A computed string-literal constant is real data and must not pass. */
const isPhantomBrandLiteral = (node: AstNode, brands: Set<string>): boolean => {
  const members = node["members"];
  if (!Array.isArray(members) || members.length === 0) {
    return false;
  }
  return members.every((member) => {
    if (
      !isAstNode(member) ||
      member.type !== "TSPropertySignature" ||
      member["computed"] !== true
    ) {
      return false;
    }
    const key = member["key"];
    return (
      isAstNode(key) &&
      key.type === "Identifier" &&
      typeof key["name"] === "string" &&
      brands.has(key["name"])
    );
  });
};

const checkIntersectionWidening = (
  node: AstNode,
  modelTypeNames: Set<string>,
  brands: Set<string>,
  context: WideningContext,
): void => {
  const members = node["types"];
  if (!Array.isArray(members)) {
    return;
  }
  let modelName: string | null = null;
  let hasDataLiteral = false;
  for (const member of members) {
    modelName ??= modelReferenceName(member, modelTypeNames);
    if (
      isAstNode(member) &&
      member.type === "TSTypeLiteral" &&
      !isPhantomBrandLiteral(member, brands)
    ) {
      hasDataLiteral = true;
    }
  }
  if (modelName === null || !hasDataLiteral || isTypePredicateAnnotation(node)) {
    return;
  }
  context.report({ node, messageId: "intersectionWidening", data: { name: modelName } });
};

// --- Rule 2: no-in-check-on-model ------------------------------------------

const FUNCTION_NODE_TYPES = new Set([
  "FunctionDeclaration",
  "FunctionExpression",
  "ArrowFunctionExpression",
]);

const SCOPE_BODY_NODE_TYPES = new Set(["BlockStatement", "Program", "StaticBlock"]);

/** The `Identifier` a binding pattern declares, unwrapping a default value. */
const bindingIdentifier = (pattern: unknown): AstNode | null => {
  if (!isAstNode(pattern)) {
    return null;
  }
  if (pattern.type === "Identifier") {
    return pattern;
  }
  if (pattern.type === "AssignmentPattern") {
    return bindingIdentifier(pattern["left"]);
  }
  return null;
};

/** The model type name the binding named `name` is annotated with in this
 * scope node (a function's params, or a block's variable declarations), or
 * `undefined` when the scope declares no such binding, or `null` when it
 * declares one without a model annotation. */
const modelAnnotationInScope = (
  scope: AstNode,
  name: string,
  modelTypeNames: Set<string>,
): string | null | undefined => {
  const candidates: unknown[] = [];
  if (FUNCTION_NODE_TYPES.has(scope.type) && Array.isArray(scope["params"])) {
    candidates.push(...scope["params"]);
  }
  if (SCOPE_BODY_NODE_TYPES.has(scope.type) && Array.isArray(scope["body"])) {
    for (const statement of scope["body"]) {
      if (isAstNode(statement) && statement.type === "VariableDeclaration") {
        const declarations = statement["declarations"];
        if (Array.isArray(declarations)) {
          for (const declaration of declarations) {
            if (isAstNode(declaration)) {
              candidates.push(declaration["id"]);
            }
          }
        }
      }
    }
  }
  for (const candidate of candidates) {
    const identifier = bindingIdentifier(candidate);
    if (identifier === null || identifier["name"] !== name) {
      continue;
    }
    const annotation = identifier["typeAnnotation"];
    if (!isAstNode(annotation) || annotation.type !== "TSTypeAnnotation") {
      return null;
    }
    return modelReferenceName(annotation["typeAnnotation"], modelTypeNames);
  }
  return undefined;
};

/** The model type the binding `name` resolves to from `from`, walking out
 * through enclosing functions and blocks to the nearest declaration. */
const modelTypeOfBinding = (
  from: AstNode,
  name: string,
  modelTypeNames: Set<string>,
): string | null => {
  let scope = from["parent"];
  while (isAstNode(scope)) {
    const found = modelAnnotationInScope(scope, name, modelTypeNames);
    if (found !== undefined) {
      return found;
    }
    scope = scope["parent"];
  }
  return null;
};

/** The root `Identifier` of `node`, unwrapping a `MemberExpression` chain
 * (`a.b.c` -> `a`), else `null`. */
const rootIdentifier = (node: unknown): AstNode | null => {
  if (!isAstNode(node)) {
    return null;
  }
  if (node.type === "Identifier") {
    return node;
  }
  if (node.type === "MemberExpression") {
    return rootIdentifier(node["object"]);
  }
  return null;
};

const sourceTextForContext = (context: InCheckContext): string => {
  if (typeof context.sourceCode?.text === "string") {
    return context.sourceCode.text;
  }
  const fallback = context.getSourceCode?.();
  return typeof fallback?.text === "string" ? fallback.text : "";
};

const textOf = (context: InCheckContext, node: AstNode): string | null => {
  const { start, end } = node;
  if (typeof start !== "number" || typeof end !== "number") {
    return null;
  }
  const text = sourceTextForContext(context);
  return text ? text.slice(start, end) : null;
};

const checkInCheckOnModel = (
  node: AstNode,
  modelTypeNames: Set<string>,
  context: InCheckContext,
): void => {
  if (node["operator"] !== "in") {
    return;
  }
  const left = node["left"];
  if (!isAstNode(left) || left.type !== "Literal" || typeof left["value"] !== "string") {
    return;
  }
  const right = node["right"];
  const root = rootIdentifier(right);
  if (root === null || typeof root["name"] !== "string") {
    return;
  }
  if (modelTypeOfBinding(node, root["name"], modelTypeNames) === null) {
    return;
  }
  const name = (isAstNode(right) ? textOf(context, right) : null) ?? root["name"];
  context.report({ node, messageId: "inCheckOnModel", data: { key: left["value"], name } });
};

export default {
  meta: { name: "folio-model-types" },
  rules: {
    "no-model-intersection-widening": {
      meta: {
        type: "problem",
        messages: {
          intersectionWidening:
            "Declare the field on the model type instead of widening `{{name}}` with a local " +
            "intersection. A field the type does not declare is invisible to every other " +
            "projection of the model, so the inverse conversion drops it. Narrow an existing " +
            "optional field with a type predicate instead. A literal whose members are all " +
            "`unique symbol` keys is a phantom brand, declares no data, and is allowed; adding " +
            "a real field to it is not.",
        },
      },
      create(context: WideningContext) {
        return {
          Program: (node: unknown) => {
            if (!isAstNode(node)) {
              return;
            }
            const modelTypeNames = collectModelTypeNames(node, context.filename);
            if (modelTypeNames.size === 0) {
              return;
            }
            const brands = uniqueSymbolNames(context);
            const visit = (child: AstNode): void => {
              if (child.type === "TSIntersectionType") {
                checkIntersectionWidening(child, modelTypeNames, brands, context);
              }
              forEachChild(child, visit);
            };
            visit(node);
          },
        };
      },
    },
    "no-in-check-on-model": {
      meta: {
        type: "problem",
        messages: {
          inCheckOnModel:
            '`"{{key}}" in {{name}}` reads a field the model type does not declare. Declare it ' +
            "on the type and read it directly; an `in` check here hides an incomplete type " +
            "from the compiler.",
        },
      },
      create(context: InCheckContext) {
        return {
          Program: (node: unknown) => {
            if (!isAstNode(node)) {
              return;
            }
            const modelTypeNames = collectModelTypeNames(node, context.filename);
            if (modelTypeNames.size === 0) {
              return;
            }
            const visit = (child: AstNode): void => {
              if (child.type === "BinaryExpression") {
                checkInCheckOnModel(child, modelTypeNames, context);
              }
              forEachChild(child, visit);
            };
            visit(node);
          },
        };
      },
    },
  },
};
