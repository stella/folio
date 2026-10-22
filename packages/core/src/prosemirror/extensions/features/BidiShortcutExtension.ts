/**
 * Bidi Shortcut Extension — Ctrl/Cmd+Left Shift → LTR, Ctrl/Cmd+Right Shift → RTL
 *
 * Uses KeyboardEvent.code to distinguish ShiftLeft vs ShiftRight, matching the
 * standard Google Docs shortcut behavior. The keydown/keyup state machine lives
 * in {@link createBidiShortcutController}; this module only wires it to the
 * editor view and dispatches the schema command on keyup.
 *
 * Priority: High (50) — should intercept before other keymaps
 */

import { Plugin } from "prosemirror-state";

// oxlint-disable-next-line import/no-cycle -- runtime-only: singletonManager consumed inside the handlers, not at module load
import { singletonManager } from "../../schema";
import { createExtension } from "../create";
import { Priority } from "../types";
import type { ExtensionContext, ExtensionRuntime } from "../types";
import { createBidiShortcutController } from "./bidiShortcutController";

export const BidiShortcutExtension = createExtension({
  name: "bidiShortcut",
  priority: Priority.High,
  onSchemaReady(_ctx: ExtensionContext): ExtensionRuntime {
    const controller = createBidiShortcutController();
    return {
      plugins: [
        new Plugin({
          props: {
            handleKeyDown(_view, event) {
              controller.handleKeyDown(event);
              // Never consume the keydown: the direction change is applied on
              // keyup, and any following key (Z, arrows, …) must reach the other
              // keymaps unimpeded.
              return false;
            },
            handleDOMEvents: {
              keyup(view, event) {
                const direction = controller.handleKeyUp(event);
                if (!direction) {
                  return false;
                }
                const cmds = singletonManager.getCommands();
                const command =
                  direction === "ltr" ? cmds["setLtr"] : cmds["setRtl"];
                command?.()(view.state, view.dispatch);
                return false;
              },
            },
          },
        }),
      ],
    };
  },
});
