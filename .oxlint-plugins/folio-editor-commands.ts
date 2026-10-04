/** Adapter commands enter the owning view's canonical journal before dispatch. */
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  filename: string;
  report: (descriptor: { node: unknown; messageId: "rawCommand" | "foreignOwner" }) => void;
};
const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" && value !== null && "type" in value && typeof value.type === "string";
const named = (node: unknown, name: string) =>
  isNode(node) && node.type === "Identifier" && node.name === name;
const member = (node: unknown, name: string) =>
  isNode(node) &&
  node.type === "MemberExpression" &&
  node.computed !== true &&
  named(node.property, name);
const containsDispatch = (node: unknown): boolean => {
  if (!isNode(node)) return false;
  if (member(node, "dispatch")) return true;
  return Object.entries(node).some(([key, value]) => {
    if (["parent", "loc", "range"].includes(key)) return false;
    return Array.isArray(value) ? value.some(containsDispatch) : containsDispatch(value);
  });
};
const OWNER = "packages/core/src/controller/hiddenEditorManager.ts";
const REACT = "packages/react/src/components/DocxEditor.tsx";
const VUE = [
  "packages/vue/src/components/Toolbar.vue",
  "packages/vue/src/composables/useFormattingActions.ts",
];
export default {
  meta: { name: "folio-editor-commands" },
  rules: {
    "command-owner-boundary": {
      meta: {
        type: "problem",
        messages: {
          rawCommand:
            "Execute adapter commands through executeEditorCommand (or executeFirstEditorCommand), so the owning view classifies canonical intents before mutation.",
          foreignOwner:
            "Only hiddenEditorManager owns registerEditorCommandOwner; adapters execute commands against the selected view.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        const react = filename.endsWith(REACT);
        const vue = VUE.some((path) => filename.endsWith(path));
        let handlerDepth = 0;
        return {
          ImportDeclaration: (node: unknown) => {
            if (!isNode(node) || filename.endsWith(OWNER)) return;
            const source = node.source;
            if (
              !isNode(source) ||
              typeof source.value !== "string" ||
              !source.value.replace(/\.(?:ts|js)$/u, "").endsWith("executeEditorCommand")
            )
              return;
            if (!Array.isArray(node.specifiers)) return;
            for (const specifier of node.specifiers) {
              if (
                isNode(specifier) &&
                (specifier.type === "ImportNamespaceSpecifier" ||
                  named(specifier.imported, "registerEditorCommandOwner"))
              )
                context.report({ node: specifier, messageId: "foreignOwner" });
            }
          },
          VariableDeclarator: (node: unknown) => {
            if (react && isNode(node) && named(node.id, "handleFormat")) handlerDepth++;
          },
          "VariableDeclarator:exit": (node: unknown) => {
            if (react && isNode(node) && named(node.id, "handleFormat")) handlerDepth--;
          },
          CallExpression: (node: unknown) => {
            if ((!vue && handlerDepth === 0) || !isNode(node) || !Array.isArray(node.arguments))
              return;
            const [state, dispatch] = node.arguments;
            if (
              (member(state, "state") || named(state, "commandState")) &&
              containsDispatch(dispatch)
            )
              context.report({ node, messageId: "rawCommand" });
          },
        };
      },
    },
  },
};
