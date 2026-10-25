/**
 * Superscript Mark Extension
 */

import { panic } from "better-result";

import { canonicalRunFormatting } from "../../canonicalCommands";
import { createMarkExtension } from "../create";
import type {
  ExtensionContext,
  ExtensionRuntime,
  MarkExtension,
} from "../types";
import {
  isMarkActive,
  toggleDocumentMark,
  withRunFormattingLock,
} from "./markUtils";

export const SuperscriptExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => MarkExtension = createMarkExtension({
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
          withRunFormattingLock(toggleDocumentMark(superscriptType), (state) =>
            canonicalRunFormatting(state, {
              vertAlign: isMarkActive(state, superscriptType)
                ? null
                : "superscript",
            }),
          ),
      },
    };
  },
});
