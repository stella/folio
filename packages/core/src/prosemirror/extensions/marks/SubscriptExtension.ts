/**
 * Subscript Mark Extension
 */

import { panic } from "better-result";

import { createMarkExtension } from "../create";
import { toggleDocumentMark } from "./markUtils";
import type { ExtensionContext, ExtensionRuntime } from "../types";

export const SubscriptExtension = createMarkExtension({
  name: "subscript",
  schemaMarkName: "subscript",
  markSpec: {
    excludes: "superscript",
    parseDOM: [{ tag: "sub" }],
    toDOM() {
      return ["sub", 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const subscriptType = ctx.schema.marks["subscript"];
    if (!subscriptType) {
      panic("Missing mark type: subscript");
    }
    return {
      commands: {
        toggleSubscript: () => toggleDocumentMark(subscriptType),
      },
    };
  },
});
