// Unrooted evaluations can lose their pending promise during collection.
// Canonical loads must use the shared transport owner.
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId: "canonicalLoadOwner";
  }) => void;
};

const FUNCTION_TYPES = new Set([
  "FunctionDeclaration",
  "FunctionExpression",
  "ArrowFunctionExpression",
]);
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);
const EXPRESSION_WRAPPERS = new Set([
  "ChainExpression",
  "TSAsExpression",
  "TSNonNullExpression",
  "TSTypeAssertion",
  "ParenthesizedExpression",
]);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const unwrap = (value: unknown): AstNode | null => {
  if (!isAstNode(value)) return null;
  if (EXPRESSION_WRAPPERS.has(value.type)) return unwrap(value["expression"]);
  return value;
};

const identifierName = (value: unknown): string | null => {
  const node = unwrap(value);
  return node?.type === "Identifier" && typeof node["name"] === "string"
    ? node["name"]
    : null;
};

const propertyName = (value: unknown): string | null => {
  const node = unwrap(value);
  if (node?.type !== "MemberExpression") return null;
  if (node["computed"] !== true) return identifierName(node["property"]);
  const property = unwrap(node["property"]);
  return property?.type === "Literal" && typeof property["value"] === "string"
    ? property["value"]
    : null;
};

const visitChildren = (
  node: AstNode,
  visit: (child: AstNode) => void,
): void => {
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) continue;
    if (Array.isArray(value)) {
      for (const item of value) if (isAstNode(item)) visit(item);
      continue;
    }
    if (isAstNode(value)) visit(value);
  }
};

type LoadAlias = "bridge" | "load";
type LoadBindings = Map<string, LoadAlias | null>;

const aliasKind = (
  value: unknown,
  aliases: ReadonlyMap<string, LoadAlias | null>,
): LoadAlias | null => {
  const node = unwrap(value);
  if (node === null) return null;
  const name = identifierName(node);
  if (name !== null) {
    if (aliases.has(name)) return aliases.get(name) ?? null;
    return name === "__folioCanonical" ? "bridge" : null;
  }
  if (node.type === "CallExpression") {
    const callee = unwrap(node["callee"]);
    if (
      propertyName(callee) === "bind" &&
      aliasKind(callee?.["object"], aliases) === "load"
    )
      return "load";
    return null;
  }
  const property = propertyName(node);
  if (property === "__folioCanonical") return "bridge";
  if (property === "load" && aliasKind(node["object"], aliases) === "bridge")
    return "load";
  return null;
};

const shadowBindings = (value: unknown, aliases: LoadBindings): void => {
  const node = unwrap(value);
  if (node === null) return;
  const name = identifierName(node);
  if (name !== null) {
    aliases.set(name, null);
    return;
  }
  if (node.type === "AssignmentPattern") {
    shadowBindings(node["left"], aliases);
    return;
  }
  if (node.type === "RestElement") {
    shadowBindings(node["argument"], aliases);
    return;
  }
  if (node.type === "Property") {
    shadowBindings(node["value"], aliases);
    return;
  }
  if (node.type === "ObjectPattern" || node.type === "ArrayPattern")
    visitChildren(node, (child) => shadowBindings(child, aliases));
};

const recordBinding = (node: AstNode, aliases: LoadBindings): void => {
  const kind = aliasKind(node["init"], aliases);
  const name = identifierName(node["id"]);
  if (name !== null) {
    aliases.set(name, kind);
    return;
  }
  shadowBindings(node["id"], aliases);
  const pattern = unwrap(node["id"]);
  if (kind !== "bridge" || pattern?.type !== "ObjectPattern") return;
  const properties = pattern["properties"];
  if (!Array.isArray(properties)) return;
  for (const property of properties) {
    if (!isAstNode(property) || property.type !== "Property") continue;
    if (identifierName(property["key"]) !== "load") continue;
    const binding = identifierName(property["value"]);
    if (binding !== null) aliases.set(binding, "load");
  }
};

const checkCanonicalLoads = (program: AstNode, context: RuleContext): void => {
  const visit = (node: AstNode, aliases: LoadBindings): void => {
    if (node.type === "AssignmentExpression") {
      const assigned = unwrap(node["right"]);
      const target = unwrap(node["left"]);
      // Installed decorators run later under the bridge owner. Exempt only the
      // function assigned to bridge.load, not arbitrary nested evaluation code.
      if (
        target?.type === "MemberExpression" &&
        propertyName(target) === "load" &&
        aliasKind(target["object"], aliases) === "bridge" &&
        assigned !== null &&
        FUNCTION_TYPES.has(assigned.type)
      ) {
        const left = unwrap(node["left"]);
        if (left !== null) visit(left, aliases);
        return;
      }
    }
    if (node.type === "BlockStatement" || FUNCTION_TYPES.has(node.type)) {
      const scope = new Map(aliases);
      if (FUNCTION_TYPES.has(node.type)) {
        shadowBindings(node["id"], scope);
        const params = node["params"];
        if (Array.isArray(params))
          for (const param of params) shadowBindings(param, scope);
      }
      visitChildren(node, (child) => visit(child, scope));
      return;
    }
    if (node.type === "VariableDeclarator") recordBinding(node, aliases);
    if (node.type === "CallExpression") {
      const callee = unwrap(node["callee"]);
      const method = propertyName(callee);
      if (
        aliasKind(callee, aliases) === "load" ||
        ((method === "call" || method === "apply") &&
          aliasKind(callee?.["object"], aliases) === "load")
      )
        context.report({ node, messageId: "canonicalLoadOwner" });
    }
    visitChildren(node, (child) => visit(child, aliases));
  };
  visit(program, new Map());
};

export default {
  meta: { name: "folio-canonical-load" },
  rules: {
    "evaluation-owner": {
      meta: {
        type: "problem",
        messages: {
          canonicalLoadOwner:
            "Use evaluateCanonicalLoad for browser-side canonical loads so the pending evaluation promise has one transport owner.",
        },
      },
      create(context: RuleContext) {
        return {
          Program: (node: unknown) => {
            if (isAstNode(node)) checkCanonicalLoads(node, context);
          },
        };
      },
    },
  },
};
