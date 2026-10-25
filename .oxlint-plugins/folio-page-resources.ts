type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId: "pageRemovalWithoutRelease" | "missingPageReleaseBoundary";
  }) => void;
};

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const nameOf = (node: unknown): string | null =>
  isNode(node) && node.type === "Identifier" && typeof node.name === "string"
    ? node.name
    : null;

const propertyOf = (node: unknown): string | null => {
  if (!isNode(node) || node.type !== "MemberExpression") return null;
  if (node.computed === true) return null;
  return nameOf(node.property);
};

const receiverOf = (node: AstNode): unknown => {
  if (node.type === "CallExpression" && isNode(node.callee))
    return node.callee.type === "MemberExpression" ? node.callee.object : null;
  if (node.type === "AssignmentExpression" && isNode(node.left))
    return node.left.type === "MemberExpression" ? node.left.object : null;
  return null;
};

const expressionPath = (node: unknown): string | null => {
  if (!isNode(node)) return null;
  if (node.type === "Identifier") return nameOf(node);
  if (node.type === "ThisExpression") return "this";
  if (node.type !== "MemberExpression" || node.computed === true) return null;
  const parent = expressionPath(node.object);
  const property = nameOf(node.property);
  return parent === null || property === null ? null : `${parent}.${property}`;
};

const positionOf = (node: AstNode): number =>
  typeof node.start === "number"
    ? node.start
    : Array.isArray(node.range) && typeof node.range[0] === "number"
      ? node.range[0]
      : Number.MAX_SAFE_INTEGER;

const descendants = (root: unknown): AstNode[] => {
  if (!isNode(root)) return [];
  const nodes: AstNode[] = [root];
  for (const [key, value] of Object.entries(root)) {
    if (
      ["parent", "loc", "range", "start", "end", "tokens", "comments"].includes(
        key,
      )
    )
      continue;
    if (Array.isArray(value)) {
      for (const child of value) nodes.push(...descendants(child));
    } else {
      nodes.push(...descendants(value));
    }
  }
  return nodes;
};

const enclosingFunction = (node: AstNode): AstNode | null => {
  let current = node.parent;
  while (isNode(current)) {
    if (
      current.type === "FunctionDeclaration" ||
      current.type === "FunctionExpression" ||
      current.type === "ArrowFunctionExpression" ||
      current.type === "MethodDefinition"
    )
      return current;
    current = current.parent;
  }
  return null;
};

const isReleaseCall = (node: AstNode): boolean =>
  node.type === "CallExpression" &&
  nameOf(node.callee) === "releaseDisplayPageResources";

const releasesMutationReceiver = (
  release: AstNode,
  mutation: AstNode,
): boolean => {
  if (!isReleaseCall(release) || !Array.isArray(release.arguments))
    return false;
  const released = expressionPath(release.arguments[0]);
  return released !== null && released === expressionPath(receiverOf(mutation));
};

const isDomRemoval = (node: AstNode): boolean => {
  if (node.type === "CallExpression") {
    if (
      propertyOf(node.callee) === "remove" &&
      propertyOf(receiverOf(node)) === "classList"
    )
      return false;
    return ["remove", "replaceWith", "replaceChildren"].includes(
      propertyOf(node.callee) ?? "",
    );
  }
  return (
    node.type === "AssignmentExpression" &&
    propertyOf(node.left) === "innerHTML"
  );
};

const expressionStatementFor = (node: AstNode): AstNode | null => {
  let current: unknown = node;
  while (isNode(current)) {
    if (current.type === "ExpressionStatement") return current;
    current = current.parent;
  }
  return null;
};

const hasAdjacentRelease = (fn: AstNode, mutation: AstNode): boolean => {
  const mutationStatement = expressionStatementFor(mutation);
  if (!mutationStatement || enclosingFunction(mutationStatement) !== fn)
    return false;
  const block = mutationStatement.parent;
  if (
    !isNode(block) ||
    block.type !== "BlockStatement" ||
    !Array.isArray(block.body)
  )
    return false;
  const statements = block.body;
  const mutationIndex = statements.indexOf(mutationStatement);
  const releaseStatement = statements.at(mutationIndex - 1);
  if (
    mutationIndex <= 0 ||
    !isNode(releaseStatement) ||
    releaseStatement.type !== "ExpressionStatement" ||
    !isNode(releaseStatement.expression) ||
    enclosingFunction(releaseStatement.expression) !== fn
  )
    return false;
  return releasesMutationReceiver(releaseStatement.expression, mutation);
};

const isTargetFile = (filename: string): boolean => {
  const normalized = filename.replaceAll("\\", "/");
  return (
    [
      "packages/core/src/layout-painter/renderPage.ts",
      "packages/core/src/layout-painter/index.ts",
      "packages/react/src/paged-editor/PagedEditor.tsx",
      "packages/vue/src/composables/useDocxEditor.ts",
    ].some((suffix) => normalized.endsWith(suffix)) ||
    /test\/__fixtures__\/page-resources\.(?:invalid|valid)\.ts$/u.test(
      normalized,
    )
  );
};

const containsRelease = (node: unknown): boolean =>
  descendants(node).some(isReleaseCall);

const hasReactCleanupBoundary = (program: AstNode): boolean =>
  descendants(program).some((node) => {
    if (
      node.type !== "CallExpression" ||
      nameOf(node.callee) !== "useLayoutEffect"
    )
      return false;
    const callback = node.arguments;
    if (!Array.isArray(callback) || !isNode(callback[0])) return false;
    return descendants(callback[0]).some(
      (entry) =>
        entry.type === "ReturnStatement" &&
        isNode(entry.argument) &&
        containsRelease(entry.argument),
    );
  });

const hasVueDestroyBoundary = (program: AstNode): boolean =>
  descendants(program).some((node) => {
    if (
      node.type !== "FunctionDeclaration" ||
      nameOf(node.id) !== "destroy" ||
      !isNode(node.body)
    )
      return false;
    return descendants(node.body).some(isReleaseCall);
  });

export default {
  meta: { name: "folio-page-resources" },
  rules: {
    "release-before-page-removal": {
      meta: {
        type: "problem",
        messages: {
          pageRemovalWithoutRelease:
            "Release rendered-page resources before removing or replacing the DOM subtree.",
          missingPageReleaseBoundary:
            "This adapter must release rendered-page resources at its teardown boundary.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        return {
          Program: (program: unknown) => {
            if (!isNode(program) || !isTargetFile(filename)) return;
            for (const node of descendants(program)) {
              if (!isDomRemoval(node)) continue;
              const fn = enclosingFunction(node);
              if (!fn || !hasAdjacentRelease(fn, node))
                context.report({
                  node,
                  messageId: "pageRemovalWithoutRelease",
                });
            }
            if (
              filename.endsWith(
                "packages/react/src/paged-editor/PagedEditor.tsx",
              ) &&
              !hasReactCleanupBoundary(program)
            )
              context.report({
                node: program,
                messageId: "missingPageReleaseBoundary",
              });
            if (
              filename.endsWith(
                "packages/vue/src/composables/useDocxEditor.ts",
              ) &&
              !hasVueDestroyBoundary(program)
            )
              context.report({
                node: program,
                messageId: "missingPageReleaseBoundary",
              });
          },
        };
      },
    },
  },
};
