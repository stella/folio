/**
 * Subscript Mark Extension
 */

import { panic } from "better-result";

import { canonicalRunFormatting } from "../../canonicalCommands";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import {
  isMarkActive,
  toggleDocumentMark,
  withRunFormattingLock,
} from "./markUtils";

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
        toggleSubscript: () =>
          withRunFormattingLock(toggleDocumentMark(subscriptType), (state) =>
            canonicalRunFormatting(state, {
              vertAlign: isMarkActive(state, subscriptType)
                ? null
                : "subscript",
            }),
          ),
      },
    };
  },
});
