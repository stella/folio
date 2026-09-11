// A ref assigned straight from render scope (`fooRef.current = foo;` as a
// direct statement of a component or hook body) mirrors a value React already
// owns. Writing that ref anywhere else is not durable: the next render
// reassigns it from state, and a setter that dedupes by identity against the
// ref sees the value as unchanged and skips the state update entirely. Route
// the change through state instead. The one legitimate exception is the setter
// that owns the ref and pairs the write with the matching state update; mark
// that site with a disable directive and a comment saying why.

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: {
    node: unknown;
    messageId: "writeToRenderMirroredRef";
    data: { name: string };
  }) => void;
};

const FUNCTION_TYPES = new Set([
  "FunctionDeclaration",
  "FunctionExpression",
  "ArrowFunctionExpression",
]);

// Keys that never lead to child nodes (or lead back up the tree).
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const identifierName = (node: unknown): string | null =>
  isAstNode(node) && node.type === "Identifier" && typeof node["name"] === "string"
    ? node["name"]
    : null;

/** Components and hooks: `Name`, `useName`. */
const isRenderScopeName = (name: string): boolean => /^(?:use[A-Z0-9_]|[A-Z])/.test(name);

/** `x.current` on the left of an assignment: the ref identifier, else null. */
const assignedRefName = (node: AstNode): string | null => {
  if (node.type !== "AssignmentExpression") {
    return null;
  }
  const left = node["left"];
  if (!isAstNode(left) || left.type !== "MemberExpression" || left["computed"] === true) {
    return null;
  }
  if (identifierName(left["property"]) !== "current") {
    return null;
  }
  return identifierName(left["object"]);
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

/** Mirror statements: `ref.current = value;` directly in the function body. */
const collectRenderMirrors = (fn: AstNode): { names: Set<string>; nodes: Set<AstNode> } => {
  const names = new Set<string>();
  const nodes = new Set<AstNode>();
  const body = fn["body"];
  if (!isAstNode(body) || body.type !== "BlockStatement" || !Array.isArray(body["body"])) {
    return { names, nodes };
  }
  for (const statement of body["body"]) {
    if (!isAstNode(statement) || statement.type !== "ExpressionStatement") {
      continue;
    }
    const expression = statement["expression"];
    if (!isAstNode(expression) || expression["operator"] !== "=") {
      continue;
    }
    const name = assignedRefName(expression);
    if (name === null) {
      continue;
    }
    names.add(name);
    nodes.add(expression);
  }
  return { names, nodes };
};

const checkRenderScope = (fn: AstNode, context: RuleContext): void => {
  const mirrors = collectRenderMirrors(fn);
  if (mirrors.names.size === 0) {
    return;
  }
  const visit = (node: AstNode): void => {
    const name = assignedRefName(node);
    if (name !== null && mirrors.names.has(name) && !mirrors.nodes.has(node)) {
      context.report({ node, messageId: "writeToRenderMirroredRef", data: { name } });
    }
    forEachChild(node, visit);
  };
  forEachChild(fn, visit);
};

/**
 * Walk the program, resolving the name a function is bound to (`function X`,
 * `const X = () =>`, `const X = memo(() =>`) so component and hook bodies can
 * be told apart from callbacks and effects nested inside them.
 */
const walk = (node: AstNode, boundName: string | null, context: RuleContext): void => {
  if (FUNCTION_TYPES.has(node.type)) {
    const ownName = identifierName(node["id"]) ?? boundName;
    if (ownName !== null && isRenderScopeName(ownName)) {
      checkRenderScope(node, context);
    }
    forEachChild(node, (child) => walk(child, null, context));
    return;
  }
  if (node.type === "VariableDeclarator") {
    const name = identifierName(node["id"]);
    const init = node["init"];
    if (isAstNode(init)) {
      walk(init, name, context);
    }
    return;
  }
  if (node.type === "CallExpression") {
    // `memo(Component)`, `forwardRef((props, ref) => ...)`: the binding name
    // carries through to the wrapped function.
    const callee = node["callee"];
    if (isAstNode(callee)) {
      walk(callee, null, context);
    }
    const args = node["arguments"];
    if (Array.isArray(args)) {
      for (const arg of args) {
        if (isAstNode(arg)) {
          walk(arg, boundName, context);
        }
      }
    }
    return;
  }
  forEachChild(node, (child) => walk(child, null, context));
};

export default {
  meta: { name: "folio-ref-mirrors" },
  rules: {
    "no-write-to-render-mirrored-ref": {
      meta: {
        type: "problem",
        messages: {
          writeToRenderMirroredRef:
            "`{{name}}.current` is reassigned from render scope on every render, so this write " +
            "is overwritten by the next render and defeats identity checks against the ref. " +
            "Update the state the ref mirrors instead. If this is the setter that owns the " +
            "ref and it also updates that state, disable this rule here and say so.",
        },
      },
      create(context: RuleContext) {
        return {
          Program: (node: unknown) => {
            if (isAstNode(node)) {
              walk(node, null, context);
            }
          },
        };
      },
    },
  },
};
