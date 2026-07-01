/**
 * Paste Cleanup Extension
 *
 * Wires two paste-handling behaviors into the editor:
 *
 * 1. `transformPastedHTML` — strips Office/web producer cruft from pasted HTML
 *    (see {@link cleanPastedHtml}) before the schema parser reads it. Runs at
 *    `Highest` priority so the cleanup happens before the style inliner and the
 *    default clipboard parser; ProseMirror chains every plugin's
 *    `transformPastedHTML`, so this cooperates with the inliner rather than
 *    replacing it, and `<style>` blocks are left intact for it to resolve.
 * 2. `Mod-Shift-v` — "paste without formatting", inserting clipboard text with
 *    the source formatting stripped (see {@link pasteWithoutFormatting}).
 */

import { Plugin } from "prosemirror-state";

import { pasteWithoutFormatting } from "../../commands/pastePlainText";
import { createExtension } from "../create";
import type { ExtensionRuntime } from "../types";
import { Priority } from "../types";
import { cleanPastedHtml } from "./pasteCleanup";

export const PasteCleanupExtension = createExtension({
  name: "pasteCleanup",
  priority: Priority.Highest,
  onSchemaReady(): ExtensionRuntime {
    const plugin = new Plugin({
      props: {
        transformPastedHTML(html: string): string {
          return cleanPastedHtml(html);
        },
      },
    });

    return {
      plugins: [plugin],
      keyboardShortcuts: {
        "Mod-Shift-v": pasteWithoutFormatting,
      },
    };
  },
});
