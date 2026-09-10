/**
 * Hard Break Extension — Shift+Enter line break
 */

import { panic } from "better-result";

import { expectHardBreakAttrs } from "../../attrs";
import { createNodeExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";

export const HardBreakExtension = createNodeExtension({
  name: "hardBreak",
  schemaNodeName: "hardBreak",
  nodeSpec: {
    inline: true,
    group: "inline",
    attrs: {
      breakType: { default: null },
      clear: { default: null },
    },
    selectable: false,
    parseDOM: [
      {
        tag: "br",
        getAttrs(node) {
          if (!(node instanceof HTMLElement)) {
            return null;
          }
          const breakType = node.dataset["docxBreakType"];
          const clear = node.dataset["docxBreakClear"];
          const attrs: Record<string, string> = {};
          if (breakType === "column" || breakType === "textWrapping") {
            attrs["breakType"] = breakType;
          }
          if (clear === "none" || clear === "left" || clear === "right" || clear === "all") {
            attrs["clear"] = clear;
          }
          return Object.keys(attrs).length > 0 ? attrs : null;
        },
      },
    ],
    toDOM(node) {
      const { breakType, clear } = expectHardBreakAttrs(node);
      const attrs: Record<string, string> = {};
      if (breakType !== undefined) {
        attrs["data-docx-break-type"] = breakType;
      }
      if (clear !== undefined) {
        attrs["data-docx-break-clear"] = clear;
      }
      return Object.keys(attrs).length > 0 ? ["br", attrs] : ["br"];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const hardBreakType = ctx.schema.nodes["hardBreak"];
    if (!hardBreakType) {
      panic("Missing node type: hardBreak");
    }

    return {
      keyboardShortcuts: {
        "Shift-Enter": (state, dispatch) => {
          if (dispatch) {
            dispatch(state.tr.replaceSelectionWith(hardBreakType.create()).scrollIntoView());
          }
          return true;
        },
      },
    };
  },
});
