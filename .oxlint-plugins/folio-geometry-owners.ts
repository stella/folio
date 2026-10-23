import { IMAGE_RESIZE_HANDLES } from "../packages/core/src/utils/imageResize";

export const IMAGE_RESIZE_ADAPTERS = {
  react: {
    file: "packages/react/src/paged-editor/ImageSelectionOverlay.tsx",
    handler: "handleResizeStart",
  },
  vue: {
    file: "packages/vue/src/components/ImageSelectionOverlay.vue",
    handler: "onResizeMove",
  },
} as const;

const OWNER_MODULE = "@stll/folio-core/utils/imageResize";
const OWNER_FUNCTION = "calculateResizedImageDimensions";
const HANDLE_VALUES = new Set<string>(IMAGE_RESIZE_HANDLES);

type AstNode = Record<string, unknown> & { type: string };
type LexicalVariable = {
  defs: readonly { node: unknown }[];
  references: readonly { writeExpr: unknown }[];
};
type LexicalScope = {
  upper: LexicalScope | null;
  set: Map<string, LexicalVariable>;
};
type MessageId = "missingOwner" | "localBounds" | "localHandles";
type RuleContext = {
  filename: string;
  sourceCode: { getScope: (node: AstNode) => LexicalScope };
  report: (descriptor: { node: unknown; messageId: MessageId }) => void;
};
type OwnedSymbol =
  "owner" | "namespace" | "math" | "bounds" | "dimensions" | "dimension";

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const identifier = (node: unknown): string | null =>
  isNode(node) && node.type === "Identifier" && typeof node.name === "string"
    ? node.name
    : null;

const literal = (node: unknown): string | null =>
  isNode(node) && node.type === "Literal" && typeof node.value === "string"
    ? node.value
    : null;

const propertyName = (node: AstNode): string | null =>
  node.computed === true ? literal(node.property) : identifier(node.property);

const descendants = (root: AstNode): AstNode[] => {
  const result: AstNode[] = [];
  const visit = (node: AstNode): void => {
    result.push(node);
    for (const [key, value] of Object.entries(node)) {
      if (key === "parent" || key === "loc" || key === "range") continue;
      if (isNode(value)) visit(value);
      else if (Array.isArray(value)) {
        for (const child of value) if (isNode(child)) visit(child);
      }
    }
  };
  visit(root);
  return result;
};

export default {
  meta: { name: "folio-geometry-owners" },
  rules: {
    "image-resize-owner": {
      meta: {
        type: "problem",
        messages: {
          missingOwner:
            "Image resize handlers must calculate dimensions with calculateResizedImageDimensions from @stll/folio-core/utils/imageResize.",
          localBounds:
            "Image dimension limits and scale clamping belong to the core imageResize owner, not framework overlays.",
          localHandles:
            "Derive image resize handles from the core ImageResizeHandle type instead of redeclaring the domain.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        const adapter = Object.values(IMAGE_RESIZE_ADAPTERS).find(
          ({ file }) => filename.endsWith(`/${file}`) || filename === file,
        );
        if (!adapter) return {};

        return {
          Program: (program: unknown) => {
            if (!isNode(program)) return;
            const nodes = descendants(program);
            const variableOf = (node: AstNode): LexicalVariable | undefined => {
              const name = identifier(node);
              if (name === null) return undefined;
              let scope: LexicalScope | null =
                context.sourceCode.getScope(node);
              while (scope) {
                const variable = scope.set.get(name);
                if (variable) return variable;
                scope = scope.upper;
              }
              return undefined;
            };
            const resolveSymbol = (
              node: unknown,
              seen = new Set<LexicalVariable>(),
            ): OwnedSymbol | undefined => {
              if (!isNode(node)) return undefined;
              const name = identifier(node);
              if (name !== null) {
                const variable = variableOf(node);
                if (!variable || variable.defs.length === 0)
                  return name === "Math" ? "math" : undefined;
                if (seen.has(variable)) return undefined;
                const nextSeen = new Set(seen);
                nextSeen.add(variable);
                const values: OwnedSymbol[] = [];
                const initializers: AstNode[] = [];
                for (const definition of variable.defs) {
                  if (!isNode(definition.node)) return undefined;
                  const declaration = definition.node;
                  if (
                    declaration.type === "ImportSpecifier" ||
                    declaration.type === "ImportNamespaceSpecifier"
                  ) {
                    const parent = declaration.parent;
                    if (
                      !isNode(parent) ||
                      parent.importKind === "type" ||
                      declaration.importKind === "type" ||
                      literal(parent.source) !== OWNER_MODULE
                    )
                      return undefined;
                    if (declaration.type === "ImportNamespaceSpecifier")
                      values.push("namespace");
                    else if (
                      identifier(declaration.imported) === OWNER_FUNCTION
                    )
                      values.push("owner");
                    else return undefined;
                    continue;
                  }
                  if (
                    declaration.type !== "VariableDeclarator" ||
                    !isNode(declaration.id)
                  )
                    return undefined;
                  if (isNode(declaration.init))
                    initializers.push(declaration.init);
                  const symbol = resolveSymbol(declaration.init, nextSeen);
                  if (!symbol) return undefined;
                  if (declaration.id.type === "Identifier") values.push(symbol);
                  else if (
                    declaration.id.type === "ObjectPattern" &&
                    Array.isArray(declaration.id.properties)
                  ) {
                    const property = declaration.id.properties.find(
                      (candidate) =>
                        isNode(candidate) &&
                        identifier(candidate.value) === name,
                    );
                    if (!isNode(property)) return undefined;
                    const key =
                      property.computed === true
                        ? literal(property.key)
                        : identifier(property.key);
                    if (symbol === "namespace" && key === OWNER_FUNCTION)
                      values.push("owner");
                    else if (
                      symbol === "math" &&
                      (key === "min" || key === "max")
                    )
                      values.push("bounds");
                    else return undefined;
                  } else return undefined;
                }
                for (const { writeExpr } of variable.references) {
                  if (!isNode(writeExpr) || initializers.includes(writeExpr))
                    continue;
                  const symbol = resolveSymbol(writeExpr, nextSeen);
                  if (!symbol) return undefined;
                  values.push(symbol);
                }
                const first = values.at(0);
                return values.every((value) => value === first)
                  ? first
                  : undefined;
              }
              if (
                node.type === "CallExpression" &&
                resolveSymbol(node.callee, seen) === "owner"
              )
                return "dimensions";
              if (node.type !== "MemberExpression") return undefined;
              const property = propertyName(node);
              const object = identifier(node.object);
              if (
                (object === "globalThis" || object === "window") &&
                property === "Math" &&
                isNode(node.object) &&
                !variableOf(node.object)?.defs.length
              )
                return "math";
              const owner = resolveSymbol(node.object, seen);
              if (
                owner === "math" &&
                (property === "min" || property === "max")
              )
                return "bounds";
              if (owner === "namespace" && property === OWNER_FUNCTION)
                return "owner";
              if (
                owner === "dimensions" &&
                (property === "width" || property === "height")
              )
                return "dimension";
              return undefined;
            };

            const handlers = nodes.filter(
              (node) =>
                (node.type === "VariableDeclarator" ||
                  node.type === "FunctionDeclaration") &&
                identifier(node.id) === adapter.handler,
            );
            if (handlers.length === 0)
              context.report({ node: program, messageId: "missingOwner" });
            for (const handler of handlers) {
              if (
                !descendants(handler).some(
                  (node) =>
                    node.type === "CallExpression" &&
                    resolveSymbol(node.callee) === "owner",
                )
              )
                context.report({ node: handler, messageId: "missingOwner" });
            }

            for (const node of nodes) {
              if (
                node.type === "MemberExpression" &&
                resolveSymbol(node) === "bounds"
              )
                context.report({ node, messageId: "localBounds" });
              if (
                (node.type === "ConditionalExpression" ||
                  node.type === "IfStatement") &&
                isNode(node.test) &&
                node.test.type === "BinaryExpression" &&
                ["<", "<=", ">", ">="].includes(String(node.test.operator))
              ) {
                const bound =
                  resolveSymbol(node.test.left) === "dimension"
                    ? node.test.right
                    : resolveSymbol(node.test.right) === "dimension"
                      ? node.test.left
                      : null;
                if (
                  isNode(bound) &&
                  bound.type === "Literal" &&
                  typeof bound.value === "number"
                ) {
                  const branches = [node.consequent, node.alternate].filter(
                    isNode,
                  );
                  if (
                    branches.some((branch) =>
                      node.type === "ConditionalExpression"
                        ? branch.type === "Literal" &&
                          branch.value === bound.value
                        : descendants(branch).some(
                            (child) =>
                              child.type === "AssignmentExpression" &&
                              resolveSymbol(child.left) === "dimension" &&
                              isNode(child.right) &&
                              child.right.type === "Literal" &&
                              child.right.value === bound.value,
                          ),
                    )
                  )
                    context.report({ node, messageId: "localBounds" });
                }
              }
              if (
                node.type === "VariableDeclarator" &&
                isNode(node.id) &&
                node.id.type === "ObjectPattern" &&
                resolveSymbol(node.init) === "math"
              ) {
                if (!Array.isArray(node.id.properties)) continue;
                for (const property of node.id.properties) {
                  if (!isNode(property)) continue;
                  const key =
                    property.computed === true
                      ? literal(property.key)
                      : identifier(property.key);
                  if (key === "min" || key === "max")
                    context.report({
                      node: property,
                      messageId: "localBounds",
                    });
                }
              }
              if (
                node.type !== "TSUnionType" ||
                !Array.isArray(node.types) ||
                node.types.length !== HANDLE_VALUES.size
              )
                continue;
              const values = node.types.map((type) =>
                isNode(type) ? literal(type.literal) : null,
              );
              if (
                new Set(values).size === HANDLE_VALUES.size &&
                values.every(
                  (value) => value !== null && HANDLE_VALUES.has(value),
                )
              )
                context.report({ node, messageId: "localHandles" });
            }
          },
        };
      },
    },
  },
};
