// Page-fragment insertion belongs to the paginator. It consumes pending
// section transitions before committing a fragment, including already-
// positioned floating and anchored content. Writing to `page.fragments`
// elsewhere bypasses that state transition and can shift section restarts.

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "directPageFragmentPush" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const memberName = (node: unknown): string | null => {
  if (!isAstNode(node) || node.type !== "MemberExpression") {
    return null;
  }
  const property = node.property;
  if (!isAstNode(property)) {
    return null;
  }
  if (property.type === "Identifier") {
    return typeof property.name === "string" ? property.name : null;
  }
  if (property.type === "Literal") {
    return typeof property.value === "string" ? property.value : null;
  }
  return null;
};

const isDirectPageFragmentPush = (node: AstNode): boolean => {
  const callee = node.callee;
  if (!isAstNode(callee) || callee.type !== "MemberExpression" || memberName(callee) !== "push") {
    return false;
  }
  const fragments = callee.object;
  if (!isAstNode(fragments) || memberName(fragments) !== "fragments") {
    return false;
  }
  return memberName(fragments.object) === "page";
};

export default {
  meta: { name: "folio-fragment-ownership" },
  rules: {
    "no-direct-page-fragment-push": {
      meta: {
        type: "problem",
        messages: {
          directPageFragmentPush:
            "Commit layout fragments through the paginator. Direct `page.fragments.push(...)` " +
            "bypasses pending section and page-number transitions; use `addFragment(...)` for " +
            "flow content or `addUnflowedFragment(...)` for already-positioned content.",
        },
      },
      create(context: RuleContext) {
        return {
          CallExpression: (node: unknown) => {
            if (isAstNode(node) && isDirectPageFragmentPush(node)) {
              context.report({ node, messageId: "directPageFragmentPush" });
            }
          },
        };
      },
    },
  },
};
