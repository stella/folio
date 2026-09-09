/**
 * Italic Mark Extension
 */

import { panic } from "better-result";
import { createMarkExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { toggleMarkForAllScripts } from "./markUtils";

export const ItalicExtension = createMarkExtension({
  name: "italic",
  schemaMarkName: "italic",
  markSpec: {
    parseDOM: [
      { tag: "i" },
      { tag: "em" },
      {
        style: "font-style",
        getAttrs: (value) => (value === "italic" ? null : false),
      },
    ],
    toDOM() {
      return ["em", 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    const italicType = ctx.schema.marks["italic"];
    if (!italicType) {
      panic("Missing mark type: italic");
    }
    const toggleItalic = toggleMarkForAllScripts(italicType, "italic");
    return {
      commands: {
        toggleItalic: () => toggleItalic,
      },
      keyboardShortcuts: {
        "Mod-i": toggleItalic,
      },
    };
  },
});
