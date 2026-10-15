/**
 * Superscript Mark Extension
 */

import { panic } from "better-result";

import { canonicalRunFormatting, withCanonicalCommand } from "../../canonicalCommands";
import { createMarkExtension } from "../create";
import { isMarkActive, toggleDocumentMark } from "./markUtils";
import type { ExtensionContext, ExtensionRuntime } from "../types";

export const SuperscriptExtension = createMarkExtension({
  name: "superscript",
  schemaMarkName: "superscript",
  markSpec: {
    excludes: "subscript",
    parseDOM: [{ tag: "sup" }],
    toDOM() {
      return ["sup", 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const superscriptType = ctx.schema.marks["superscript"];
    if (!superscriptType) {
      panic("Missing mark type: superscript");
    }
    return {
      commands: {
        toggleSuperscript: () =>
          withCanonicalCommand(toggleDocumentMark(superscriptType), (state) =>
            canonicalRunFormatting(state, {
              vertAlign: isMarkActive(state, superscriptType) ? null : "superscript",
            }),
          ),
      },
    };
  },
});
