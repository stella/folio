/** Fresh range deletions must preserve revisions on already deleted runs. */
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  filename: string;
  report: (descriptor: { node: unknown; messageId: "rawDeletion" }) => void;
};

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" && value !== null && "type" in value && typeof value.type === "string";
const nameOf = (node: unknown): string | undefined => {
  if (!isNode(node)) return undefined;
  if (node.type === "Identifier" && typeof node.name === "string") return node.name;
  if (node.type === "Literal" && typeof node.value === "string") return node.value;
  return undefined;
};
const member = (node: unknown, name: string): node is AstNode =>
  isNode(node) && node.type === "MemberExpression" && nameOf(node.property) === name;

export default {
  meta: { name: "folio-deletion-marks" },
  rules: {
    "preserve-pending-deletions": {
      meta: {
        type: "problem",
        messages: {
          rawDeletion:
            "Apply fresh deletions through addTrackedDeletionMark, which preserves pending deletion ownership and ancestry.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        if (
          !filename.includes("packages/core/src/") ||
          filename.endsWith(".test.ts") ||
          filename.endsWith("/addTrackedDeletionMark.ts")
        )
          return {};
        const bindings = new Map<string, unknown[]>();
        const calls: AstNode[] = [];
        const resolves = (
          node: unknown,
          predicate: (value: AstNode) => boolean,
          seen = new Set<string>(),
        ): boolean => {
          if (!isNode(node)) return false;
          if (predicate(node)) return true;
          if (node.type === "Identifier" && typeof node.name === "string") {
            if (seen.has(node.name)) return false;
            const next = new Set(seen).add(node.name);
            return (bindings.get(node.name) ?? []).some((value) =>
              resolves(value, predicate, next),
            );
          }
          if (node.type === "LogicalExpression")
            return resolves(node.left, predicate, seen) || resolves(node.right, predicate, seen);
          if (node.type === "ConditionalExpression")
            return (
              resolves(node.consequent, predicate, seen) ||
              resolves(node.alternate, predicate, seen)
            );
          return false;
        };
        const deletionType = (node: AstNode) =>
          member(node, "deletion") && member(node.object, "marks");
        const deletionMark = (node: AstNode): boolean => {
          if (node.type !== "CallExpression" || !isNode(node.callee)) return false;
          if (member(node.callee, "create")) return resolves(node.callee.object, deletionType);
          return (
            member(node.callee, "mark") &&
            Array.isArray(node.arguments) &&
            nameOf(node.arguments.at(0)) === "deletion"
          );
        };
        return {
          AssignmentExpression: (node: unknown) => {
            if (
              !isNode(node) ||
              !isNode(node.left) ||
              node.left.type !== "Identifier" ||
              typeof node.left.name !== "string"
            )
              return;
            const values = bindings.get(node.left.name) ?? [];
            values.push(node.right);
            bindings.set(node.left.name, values);
          },
          VariableDeclarator: (node: unknown) => {
            if (!isNode(node) || !isNode(node.id)) return;
            if (node.id.type === "ObjectPattern" && Array.isArray(node.id.properties)) {
              for (const property of node.id.properties) {
                if (
                  !isNode(property) ||
                  !isNode(property.value) ||
                  property.value.type !== "Identifier" ||
                  typeof property.value.name !== "string"
                )
                  continue;
                const values = bindings.get(property.value.name) ?? [];
                values.push({
                  type: "MemberExpression",
                  object: node.init,
                  property: property.key,
                });
                bindings.set(property.value.name, values);
              }
              return;
            }
            if (node.id.type !== "Identifier" || typeof node.id.name !== "string") return;
            const values = bindings.get(node.id.name) ?? [];
            values.push(node.init);
            bindings.set(node.id.name, values);
          },
          CallExpression: (node: unknown) => {
            if (isNode(node) && member(node.callee, "addMark")) calls.push(node);
          },
          "Program:exit": () => {
            for (const call of calls) {
              if (Array.isArray(call.arguments) && resolves(call.arguments.at(2), deletionMark))
                context.report({ node: call, messageId: "rawDeletion" });
            }
          },
        };
      },
    },
  },
};
