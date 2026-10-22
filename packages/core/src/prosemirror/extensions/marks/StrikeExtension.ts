/**
 * Strikethrough Mark Extension
 */

import { panic } from "better-result";

import {
  canonicalRunFormatting,
  withCanonicalCommand,
} from "../../canonicalCommands";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { isMarkActive, toggleDocumentMark } from "./markUtils";

export const StrikeExtension = createMarkExtension({
  name: "strike",
  schemaMarkName: "strike",
  markSpec: {
    attrs: {
      double: { default: false },
    },
    parseDOM: [
      { tag: "s" },
      { tag: "strike" },
      { tag: "del" },
      {
        style: "text-decoration",
        getAttrs: (value) => (value.includes("line-through") ? {} : false),
      },
    ],
    toDOM() {
      return ["s", 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const strikeType = ctx.schema.marks["strike"];
    if (!strikeType) {
      panic("Missing mark type: strike");
    }
    return {
      commands: {
        toggleStrike: () =>
          withCanonicalCommand(toggleDocumentMark(strikeType), (state) => {
            const active = isMarkActive(state, strikeType);
            return canonicalRunFormatting(state, {
              strike: active ? null : true,
              doubleStrike: null,
            });
          }),
      },
    };
  },
});
