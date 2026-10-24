// Source-backed inline runs carry metadata in ProseMirror marks. Rebuilt runs
// must restore the run record before leaving their control-flow region.

type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId: "unrestoredSourceRun";
  }) => void;
};

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const children = (node: AstNode): AstNode[] => {
  const result: AstNode[] = [];
  for (const [key, value] of Object.entries(node)) {
    if (["parent", "loc", "range", "tokens", "comments"].includes(key))
      continue;
    if (Array.isArray(value)) {
      for (const item of value) if (isNode(item)) result.push(item);
    } else if (isNode(value)) {
      result.push(value);
    }
  }
  return result;
};

const identifier = (node: unknown): string | undefined =>
  isNode(node) && node.type === "Identifier" && typeof node.name === "string"
    ? node.name
    : undefined;

const propertyName = (node: unknown): string | undefined => {
  if (!isNode(node)) return undefined;
  if (node.type === "Identifier" && typeof node.name === "string")
    return node.name;
  if (node.type === "Literal" && typeof node.value === "string")
    return node.value;
  return undefined;
};

const objectProperty = (node: AstNode, name: string): unknown => {
  if (node.type !== "ObjectExpression" || !Array.isArray(node.properties))
    return undefined;
  for (const property of node.properties) {
    if (!isNode(property) || property.type !== "Property") continue;
    if (propertyName(property.key) === name) return property.value;
  }
  return undefined;
};

const isFunction = (node: AstNode): boolean =>
  node.type === "FunctionDeclaration" ||
  node.type === "FunctionExpression" ||
  node.type === "ArrowFunctionExpression";

const bindingNames = (parameter: unknown): string[] => {
  const names: string[] = [];
  const visit = (node: unknown): void => {
    if (!isNode(node)) return;
    if (node.type === "Identifier") {
      if (typeof node.name === "string") names.push(node.name);
      return;
    }
    if (node.type === "AssignmentPattern") {
      visit(node.left);
      return;
    }
    if (node.type === "RestElement" || node.type === "TSParameterProperty") {
      visit(node.argument ?? node.parameter);
      return;
    }
    if (node.type === "ArrayPattern") {
      for (const element of Array.isArray(node.elements) ? node.elements : []) {
        visit(element);
      }
      return;
    }
    if (node.type === "ObjectPattern") {
      for (const property of Array.isArray(node.properties)
        ? node.properties
        : []) {
        if (!isNode(property)) continue;
        visit(
          property.type === "RestElement" ? property.argument : property.value,
        );
      }
    }
  };
  visit(parameter);
  return names;
};

type TreeIndex = {
  nodesByFunction: Map<AstNode, AstNode[]>;
  parent: WeakMap<object, AstNode>;
};

const indexTree = (root: AstNode): TreeIndex => {
  const nodesByFunction = new Map<AstNode, AstNode[]>();
  const parent = new WeakMap<object, AstNode>();
  const seen = new WeakSet<object>();
  const visit = (
    node: AstNode,
    owner?: AstNode,
    parentNode?: AstNode,
  ): void => {
    if (seen.has(node)) return;
    seen.add(node);
    if (parentNode) parent.set(node, parentNode);
    const nextOwner = isFunction(node) ? node : owner;
    if (nextOwner) {
      const owned = nodesByFunction.get(nextOwner) ?? [];
      owned.push(node);
      nodesByFunction.set(nextOwner, owned);
    }
    for (const child of children(node)) visit(child, nextOwner, node);
  };
  visit(root);
  return { nodesByFunction, parent };
};

const ancestorsUntil = (
  node: AstNode,
  stop: AstNode,
  parent: WeakMap<object, AstNode>,
): AstNode[] => {
  const result: AstNode[] = [];
  let current = parent.get(node);
  while (current && current !== stop) {
    result.push(current);
    current = parent.get(current);
  }
  return result;
};

const listOwner = (
  node: AstNode,
  fn: AstNode,
  parent: WeakMap<object, AstNode>,
): AstNode | undefined =>
  ancestorsUntil(node, fn, parent).find(
    (ancestor) =>
      ancestor.type === "BlockStatement" || ancestor.type === "SwitchCase",
  );

const directStatement = (
  node: AstNode,
  owner: AstNode,
  parent: WeakMap<object, AstNode>,
): AstNode | undefined => {
  let current = node;
  let next = parent.get(current);
  while (next && next !== owner) {
    current = next;
    next = parent.get(current);
  }
  return next === owner ? current : undefined;
};

const statementList = (owner: AstNode): AstNode[] => {
  const items = owner.type === "SwitchCase" ? owner.consequent : owner.body;
  return Array.isArray(items) ? items.filter(isNode) : [];
};

const containsExit = (node: AstNode): boolean =>
  node.type === "ReturnStatement" ||
  node.type === "ThrowStatement" ||
  (!isFunction(node) && children(node).some(containsExit));

const exitsBeforeRestoration = (
  run: AstNode,
  call: AstNode,
  owner: AstNode | undefined,
  parent: WeakMap<object, AstNode>,
): boolean => {
  if (!owner) return true;
  const runStatement = directStatement(run, owner, parent);
  const callStatement = directStatement(call, owner, parent);
  if (!runStatement || !callStatement) return true;
  const statements = statementList(owner);
  const runIndex = statements.indexOf(runStatement);
  const callIndex = statements.indexOf(callStatement);
  if (runIndex < 0 || callIndex <= runIndex) return true;
  return statements.slice(runIndex + 1, callIndex).some(containsExit);
};

const hasShadowingBinding = (
  call: AstNode,
  owner: AstNode | undefined,
  runName: string,
  parent: WeakMap<object, AstNode>,
): boolean => {
  let current = parent.get(call);
  while (current && current !== owner) {
    if (current.type === "BlockStatement") {
      for (const statement of statementList(current)) {
        if (
          statement.type === "VariableDeclaration" &&
          Array.isArray(statement.declarations) &&
          statement.declarations.some(
            (declaration) =>
              isNode(declaration) && identifier(declaration.id) === runName,
          )
        ) {
          return true;
        }
      }
    }
    current = parent.get(current);
  }
  return false;
};

const contains = (root: AstNode, target: AstNode): boolean =>
  root === target || children(root).some((child) => contains(child, target));

const isSourceMarks = (node: unknown, sources: Set<string>): boolean => {
  const name = identifier(node);
  if (name && sources.has(name)) return true;
  if (
    isNode(node) &&
    node.type === "MemberExpression" &&
    propertyName(node.property) === "marks"
  ) {
    const objectName = identifier(node.object);
    return objectName !== undefined && sources.has(objectName + ".marks");
  }
  return false;
};

const findRunVariable = (
  run: AstNode,
  parent: WeakMap<object, AstNode>,
): string | undefined => {
  const owner = parent.get(run);
  if (!owner) return undefined;
  if (owner.type === "VariableDeclarator" && owner.init === run) {
    return identifier(owner.id);
  }
  if (owner.type === "AssignmentExpression" && owner.right === run) {
    return identifier(owner.left);
  }
  return undefined;
};

const restores = (
  call: AstNode,
  runName: string,
  sources: Set<string>,
): boolean => {
  if (
    call.type !== "CallExpression" ||
    !isNode(call.callee) ||
    identifier(call.callee) !== "restoreRunRecord"
  ) {
    return false;
  }
  const args = Array.isArray(call.arguments) ? call.arguments : [];
  return (
    identifier(args.at(0)) === runName && isSourceMarks(args.at(1), sources)
  );
};

const marksGuardedCall = (
  call: AstNode,
  fn: AstNode,
  runOwner: AstNode | undefined,
  runName: string,
  sources: Set<string>,
  parent: WeakMap<object, AstNode>,
): boolean => {
  for (const ancestor of ancestorsUntil(call, fn, parent)) {
    if (
      ancestor.type !== "IfStatement" ||
      !isNode(ancestor.test) ||
      !isNode(ancestor.consequent)
    ) {
      continue;
    }
    if (
      isSourceMarks(ancestor.test, sources) &&
      contains(ancestor.consequent, call) &&
      listOwner(ancestor, fn, parent) === runOwner &&
      ancestorsUntil(call, ancestor, parent).every(
        (between) =>
          between.type === "BlockStatement" ||
          between.type === "ExpressionStatement",
      )
    ) {
      return restores(call, runName, sources);
    }
  }
  return false;
};

const hasRestoration = (
  run: AstNode,
  fn: AstNode,
  sources: Set<string>,
  parent: WeakMap<object, AstNode>,
  calls: AstNode[],
): boolean => {
  const runName = findRunVariable(run, parent);
  if (!runName) return false;
  const owner = listOwner(run, fn, parent);
  const runEnd = typeof run.end === "number" ? run.end : undefined;
  return calls.some((call) => {
    if (!restores(call, runName, sources)) return false;
    const callStart = typeof call.start === "number" ? call.start : undefined;
    if (runEnd !== undefined && callStart !== undefined && callStart < runEnd) {
      return false;
    }
    return (
      !hasShadowingBinding(call, owner, runName, parent) &&
      !exitsBeforeRestoration(run, call, owner, parent) &&
      (listOwner(call, fn, parent) === owner ||
        marksGuardedCall(call, fn, owner, runName, sources, parent))
    );
  });
};

const analyzeFunction = (
  fn: AstNode,
  nodes: AstNode[],
  parent: WeakMap<object, AstNode>,
  report: RuleContext["report"],
): void => {
  const params = Array.isArray(fn.params) ? fn.params : [];
  const names = new Set(params.flatMap(bindingNames));
  const sources = new Set<string>();
  if (names.has("marks")) sources.add("marks");
  for (const node of nodes) {
    if (
      node.type === "MemberExpression" &&
      propertyName(node.property) === "marks"
    ) {
      const objectName = identifier(node.object);
      if (objectName && names.has(objectName)) {
        sources.add(objectName + ".marks");
      }
    }
  }
  if (sources.size === 0) return;

  const calls = nodes.filter(
    (node) =>
      node.type === "CallExpression" &&
      isNode(node.callee) &&
      identifier(node.callee) === "restoreRunRecord",
  );
  const runs = nodes.filter(
    (node) =>
      node.type === "ObjectExpression" &&
      propertyName(objectProperty(node, "type")) === "run",
  );
  for (const run of runs) {
    if (!hasRestoration(run, fn, sources, parent, calls)) {
      report({ node: run, messageId: "unrestoredSourceRun" });
    }
  }
};

export default {
  meta: { name: "folio-run-record" },
  rules: {
    "restore-source-run-record": {
      meta: {
        type: "problem",
        messages: {
          unrestoredSourceRun:
            "Restore this source-backed Run with restoreRunRecord(run, marks) in the same control-flow region.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        if (
          !filename.endsWith(
            "packages/core/src/prosemirror/conversion/fromProseDoc.ts",
          ) &&
          !filename.endsWith("run-record-metadata.invalid.ts") &&
          !filename.endsWith("run-record-metadata.valid.ts")
        ) {
          return {};
        }
        return {
          Program: (node: unknown) => {
            if (!isNode(node)) return;
            const index = indexTree(node);
            for (const [fn, nodes] of index.nodesByFunction) {
              analyzeFunction(fn, nodes, index.parent, context.report);
            }
          },
        };
      },
    },
  },
};
