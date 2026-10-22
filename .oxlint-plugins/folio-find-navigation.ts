/** Search navigation belongs to the shared view, including page and thread activation. */
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "navigationOwner" }) => void;
};
const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end"]);
const WRAPPERS = new Set([
  "ChainExpression",
  "TSAsExpression",
  "TSNonNullExpression",
]);
const unwrap = (value: unknown): AstNode | null => {
  if (!isNode(value)) return null;
  return WRAPPERS.has(value.type) ? unwrap(value.expression) : value;
};
const name = (value: unknown): string | null => {
  const node = unwrap(value);
  if (node?.type === "Identifier" && typeof node.name === "string")
    return node.name;
  if (node?.type === "Literal" && typeof node.value === "string")
    return node.value;
  return null;
};
const member = (value: unknown): string | null => {
  const node = unwrap(value);
  if (node?.type !== "MemberExpression") return null;
  const property = unwrap(node.property);
  if (node.computed === true && property?.type !== "Literal") return null;
  return name(property);
};
const children = (node: AstNode): AstNode[] => {
  const result: AstNode[] = [];
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) continue;
    if (Array.isArray(value)) {
      for (const item of value) if (isNode(item)) result.push(item);
    } else if (isNode(value)) result.push(value);
  }
  return result;
};
const rootName = (value: unknown): string | null => {
  const node = unwrap(value);
  if (node?.type === "MemberExpression") return rootName(node.object);
  return name(node);
};
const check = (program: AstNode, context: RuleContext): void => {
  const nodes: AstNode[] = [];
  const collect = (node: AstNode): void => {
    nodes.push(node);
    for (const child of children(node)) collect(child);
  };
  collect(program);
  const factories = new Set<string>();
  const namespaces = new Set<string>();
  const pageFallbacks = new Set<string>();
  for (const node of nodes) {
    if (node.type !== "ImportDeclaration") continue;
    const source = unwrap(node.source);
    if (typeof source?.value !== "string") continue;
    const specifiers = node.specifiers;
    if (!Array.isArray(specifiers)) continue;
    for (const specifier of specifiers) {
      if (!isNode(specifier)) continue;
      const local = name(specifier.local);
      if (local === null) continue;
      if (
        source.value.replace(/\.(?:ts|js)$/u, "").endsWith("/findMatchView") &&
        specifier.type === "ImportNamespaceSpecifier"
      )
        namespaces.add(local);
      if (
        source.value.replace(/\.(?:ts|js)$/u, "").endsWith("/findMatchView") &&
        name(specifier.imported) === "createFindMatchView"
      )
        factories.add(local);
      if (name(specifier.imported) === "scrollToMatch")
        pageFallbacks.add(local);
    }
  }
  if (factories.size === 0 && namespaces.size === 0) return;
  const isFactory = (value: unknown): boolean => {
    const node = unwrap(value);
    return (
      factories.has(name(node) ?? "") ||
      (node?.type === "MemberExpression" &&
        member(node) === "createFindMatchView" &&
        namespaces.has(name(node.object) ?? "")) ||
      (node?.type === "TSQualifiedName" &&
        name(node.right) === "createFindMatchView" &&
        namespaces.has(name(node.left) ?? ""))
    );
  };
  const views = new Set<string>();
  const producesView = (value: unknown): boolean => {
    const node = unwrap(value);
    if (node === null) return false;
    if (node.type === "CallExpression") {
      if (isFactory(node.callee)) return true;
      return (
        producesView(node.typeArguments) || producesView(node.typeParameters)
      );
    }
    if (node.type === "TSTypeQuery" && isFactory(node.exprName)) return true;
    if (node.type === "Identifier") return views.has(name(node) ?? "");
    if (node.type === "MemberExpression")
      return views.has(rootName(node) ?? "");
    if (
      [
        "ConditionalExpression",
        "LogicalExpression",
        "TSTypeAnnotation",
        "TSUnionType",
        "TSTypeReference",
        "TSTypeParameterInstantiation",
      ].includes(node.type)
    )
      return children(node).some(producesView);
    return false;
  };
  // Assignments often occur after the navigation callback; resolve aliases
  // before checking uses, rather than depending on declaration order.
  let previousSize = -1;
  while (previousSize !== views.size) {
    previousSize = views.size;
    for (const node of nodes) {
      if (
        node.type === "VariableDeclarator" &&
        (producesView(node.init) ||
          (isNode(node.id) && producesView(node.id.typeAnnotation)))
      ) {
        const binding = name(node.id);
        if (binding !== null) views.add(binding);
      }
      if (node.type === "AssignmentExpression" && producesView(node.right)) {
        const binding = rootName(node.left);
        if (binding !== null) views.add(binding);
      }
    }
  }
  const selectsPage = (value: unknown): boolean => {
    const node = unwrap(value);
    if (node === null) return false;
    if (node.type === "Literal" && typeof node.value === "string")
      return /\[data-page-(?:number|index)/u.test(node.value);
    if (node.type === "TemplateLiteral")
      return children(node).some(selectsPage);
    if (
      node.type !== "TemplateElement" ||
      typeof node.value !== "object" ||
      node.value === null
    )
      return false;
    return (
      "raw" in node.value &&
      typeof node.value.raw === "string" &&
      /\[data-page-(?:number|index)/u.test(node.value.raw)
    );
  };
  for (const node of nodes) {
    if (
      node.type === "MemberExpression" &&
      member(node) === "reveal" &&
      producesView(node.object)
    )
      context.report({ node, messageId: "navigationOwner" });
    if (
      node.type === "CallExpression" &&
      pageFallbacks.has(name(node.callee) ?? "")
    )
      context.report({ node, messageId: "navigationOwner" });
    if (
      node.type === "CallExpression" &&
      ["querySelector", "querySelectorAll"].includes(
        member(node.callee) ?? "",
      ) &&
      Array.isArray(node.arguments) &&
      node.arguments.some(selectsPage)
    )
      context.report({ node, messageId: "navigationOwner" });
    if (node.type !== "VariableDeclarator" || !producesView(node.init))
      continue;
    const pattern = unwrap(node.id);
    if (pattern?.type !== "ObjectPattern") continue;
    for (const property of children(pattern))
      if (property.type === "Property" && name(property.key) === "reveal")
        context.report({ node: property, messageId: "navigationOwner" });
  }
};
export default {
  meta: { name: "folio-find-navigation" },
  rules: {
    "navigation-owner": {
      meta: {
        type: "problem",
        messages: {
          navigationOwner:
            "Navigate search matches through FindMatchView.navigate so page mounting and comment activation have one owner.",
        },
      },
      create(context: RuleContext) {
        return {
          Program: (node: unknown) => {
            if (isNode(node)) check(node, context);
          },
        };
      },
    },
  },
};
