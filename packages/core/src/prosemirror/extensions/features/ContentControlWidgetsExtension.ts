/**
 * Content-Control Widgets Extension — wires the interactive widget plugin
 * into the editor so clicking inside a typed content control invokes the
 * appropriate transaction.
 *
 * The plugin owns checkbox toggling outright; dropdown / date picker chrome
 * is delegated to the editor shell via the `onWidgetEvent` callback (the
 * shell renders the menu / picker and dispatches via `dispatchDropdownPick`
 * / `dispatchDatePick`).
 */

import { createContentControlWidgetsPlugin } from "../../plugins/contentControlWidgets";
import { createExtension } from "../create";
import type { Extension, ExtensionContext, ExtensionRuntime } from "../types";

export const ContentControlWidgetsExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => Extension = createExtension({
  name: "contentControlWidgets",
  onSchemaReady(_ctx: ExtensionContext): ExtensionRuntime {
    return {
      plugins: [createContentControlWidgetsPlugin()],
    };
  },
});
