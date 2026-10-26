/**
 * Selection Tracker Extension — wraps createSelectionTrackerPlugin
 */

import { createSelectionTrackerPlugin } from "../../plugins/selectionTracker";
import type { SelectionChangeCallback } from "../../plugins/selectionTracker";
import { createExtension } from "../create";
import type { Extension, ExtensionContext, ExtensionRuntime } from "../types";

type SelectionTrackerOptions = {
  onSelectionChange?: SelectionChangeCallback;
};

const defaultSelectionTrackerOptions: SelectionTrackerOptions = {};

export const SelectionTrackerExtension: (
  options?: Partial<SelectionTrackerOptions> | undefined,
) => Extension = createExtension({
  name: "selectionTracker",
  defaultOptions: defaultSelectionTrackerOptions,
  onSchemaReady(
    _ctx: ExtensionContext,
    options: SelectionTrackerOptions,
  ): ExtensionRuntime {
    return {
      plugins: [createSelectionTrackerPlugin(options.onSelectionChange)],
    };
  },
});
