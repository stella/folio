import { RUN_FORMATTING_MARK_NAMES } from "../packages/core/src/prosemirror/runFormattingMarkNames.ts";

/** Adapter commands enter the owning view's canonical journal before dispatch. */
type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId:
      | "rawCommand"
      | "foreignOwner"
      | "rawRunFormattingCommand"
      | "resizeOwner";
  }) => void;
};
const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";
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
    return Array.isArray(value)
      ? value.some(containsDispatch)
      : containsDispatch(value);
  });
};
const visit = (node: unknown, callback: (node: AstNode) => void): void => {
  if (!isNode(node)) return;
  callback(node);
  for (const [key, value] of Object.entries(node)) {
    if (["parent", "loc", "range"].includes(key)) continue;
    if (Array.isArray(value)) value.forEach((child) => visit(child, callback));
    else visit(value, callback);
  }
};
const stringValue = (node: unknown): string | undefined =>
  !isNode(node)
    ? undefined
    : node.type === "Literal" && typeof node.value === "string"
      ? node.value
      : node.type === "Identifier" && typeof node.name === "string"
        ? node.name
        : undefined;
const OWNER = "packages/core/src/controller/hiddenEditorManager.ts";
const REACT = "packages/react/src/components/DocxEditor.tsx";
const VUE = [
  "packages/vue/src/components/Toolbar.vue",
  "packages/vue/src/composables/useFormattingActions.ts",
];
export default {
  meta: { name: "folio-editor-commands" },
  rules: {
    "table-resize-owner": {
      meta: {
        type: "problem",
        messages: {
          resizeOwner:
            "Commit table resize gestures through core tableResize helpers so every adapter preserves the same grid and spanning-cell widths.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        const react = filename.endsWith(
          "packages/react/src/paged-editor/PagedEditor.tsx",
        );
        const vue = filename.endsWith(
          "packages/vue/src/composables/useTableResize.ts",
        );
        let resizeDepth = 0;
        const entersResize = (node: unknown) =>
          isNode(node) &&
          ((react && named(node.id, "handleMouseUp")) ||
            (vue && named(node.id, "handleUp")));
        const enter = (node: unknown) => {
          if (entersResize(node)) resizeDepth++;
        };
        const leave = (node: unknown) => {
          if (entersResize(node)) resizeDepth--;
        };
        return {
          VariableDeclarator: enter,
          "VariableDeclarator:exit": leave,
          FunctionDeclaration: enter,
          "FunctionDeclaration:exit": leave,
          CallExpression: (node: unknown) => {
            if (resizeDepth === 0 || !isNode(node) || !isNode(node.callee))
              return;
            const callee = node.callee;
            if (callee.type !== "MemberExpression") return;
            const property = callee.property;
            if (
              named(property, "setNodeMarkup") ||
              (isNode(property) &&
                property.type === "Literal" &&
                property.value === "setNodeMarkup")
            ) {
              context.report({ node, messageId: "resizeOwner" });
            }
          },
        };
      },
    },
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
              !source.value
                .replace(/\.(?:ts|js)$/u, "")
                .endsWith("executeEditorCommand")
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
            if (react && isNode(node) && named(node.id, "handleFormat"))
              handlerDepth++;
          },
          "VariableDeclarator:exit": (node: unknown) => {
            if (react && isNode(node) && named(node.id, "handleFormat"))
              handlerDepth--;
          },
          CallExpression: (node: unknown) => {
            if (
              (!vue && handlerDepth === 0) ||
              !isNode(node) ||
              !Array.isArray(node.arguments)
            )
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
    "run-formatting-lock-boundary": {
      meta: {
        type: "problem",
        messages: {
          rawRunFormattingCommand:
            "Run-formatting commands must use withRunFormattingLock so direct execution and canonical intent planning share content-lock policy.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        if (
          !filename.includes(
            "/packages/core/src/prosemirror/extensions/marks/",
          ) ||
          !filename.endsWith("Extension.ts")
        )
          return {};
        return {
          Program: (program: unknown) => {
            const bindings = new Set<string>();
            const markNames = new Set<string>();
            const calls: AstNode[] = [];
            visit(program, (node) => {
              if (
                node.type === "ImportDeclaration" &&
                Array.isArray(node.specifiers)
              ) {
                for (const specifier of node.specifiers) {
                  if (
                    isNode(specifier) &&
                    specifier.type === "ImportSpecifier" &&
                    stringValue(specifier.imported) ===
                      "withCanonicalCommand" &&
                    isNode(specifier.local) &&
                    typeof specifier.local.name === "string"
                  )
                    bindings.add(specifier.local.name);
                }
              }
              if (
                (node.type === "Property" ||
                  node.type === "PropertyDefinition") &&
                isNode(node.key) &&
                node.key.name === "schemaMarkName"
              ) {
                const markName = stringValue(node.value);
                if (markName) markNames.add(markName);
              }
            });
            visit(program, (node) => {
              if (
                node.type === "CallExpression" &&
                isNode(node.callee) &&
                node.callee.type === "Identifier" &&
                typeof node.callee.name === "string" &&
                bindings.has(node.callee.name)
              )
                calls.push(node);
            });
            if (
              ![...markNames].some((name) =>
                RUN_FORMATTING_MARK_NAMES.has(name),
              )
            )
              return;
            for (const node of calls)
              context.report({ node, messageId: "rawRunFormattingCommand" });
          },
        };
      },
    },
  },
};
