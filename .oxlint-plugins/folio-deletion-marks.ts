/** Fresh range deletions must preserve revisions on already deleted runs. */
type AstNode = Record<string, unknown> & { type: string };
type LexicalVariable = {
  defs: readonly { node: unknown }[];
  references: readonly { writeExpr: unknown }[];
};
type LexicalScope = {
  upper: LexicalScope | null;
  set: Map<string, LexicalVariable>;
};
type RuleContext = {
  filename: string;
  sourceCode: { getScope: (node: AstNode) => LexicalScope };
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
        const variableOf = (node: AstNode): LexicalVariable | undefined => {
          if (typeof node.name !== "string") return undefined;
          let scope: LexicalScope | null = context.sourceCode.getScope(node);
          while (scope) {
            const variable = scope.set.get(node.name);
            if (variable) return variable;
            scope = scope.upper;
          }
          return undefined;
        };
        const valuesOf = (variable: LexicalVariable): unknown[] => {
          const values = variable.references.map(({ writeExpr }) => writeExpr);
          for (const { node } of variable.defs) {
            if (!isNode(node) || node.type !== "VariableDeclarator" || !isNode(node.id)) continue;
            if (node.id.type === "Identifier") values.push(node.init);
            if (node.id.type !== "ObjectPattern" || !Array.isArray(node.id.properties)) continue;
            for (const property of node.id.properties) {
              if (
                !isNode(property) ||
                !isNode(property.value) ||
                property.value.type !== "Identifier"
              )
                continue;
              if (variableOf(property.value) !== variable) continue;
              values.push({ type: "MemberExpression", object: node.init, property: property.key });
            }
          }
          return values;
        };
        const calls: AstNode[] = [];
        const resolves = (
          node: unknown,
          predicate: (value: AstNode) => boolean,
          seen = new Set<LexicalVariable>(),
        ): boolean => {
          if (!isNode(node)) return false;
          if (predicate(node)) return true;
          if (node.type === "Identifier" && typeof node.name === "string") {
            const variable = variableOf(node);
            if (!variable || seen.has(variable)) return false;
            const next = new Set(seen).add(variable);
            return valuesOf(variable).some((value) => resolves(value, predicate, next));
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
