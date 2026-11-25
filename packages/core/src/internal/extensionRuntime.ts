import { keymap } from "prosemirror-keymap";
import type { Schema } from "prosemirror-model";
import type { Plugin as PMPlugin } from "prosemirror-state";

import type {
  AnyExtension,
  ExtensionCommandMap,
  ExtensionContext,
  KeyboardShortcutMap,
} from "../prosemirror/extensions/types";
import { createNativeNavigationPlugin } from "../prosemirror/nativeSelection";
import {
  areTransactionInvariantsEnabled,
  createTransactionInvariantPlugin,
} from "../prosemirror/plugins/transactionInvariants";
import { createTextInputPlugin } from "../prosemirror/textInput";

export const createExtensionRuntime = (
  extensions: readonly AnyExtension[],
  schema: Schema,
) => {
  const ctx: ExtensionContext = { schema };
  const allKeyboardShortcuts: KeyboardShortcutMap[] = [];
  const allPlugins: PMPlugin[] = [];
  const allCommands: ExtensionCommandMap = {};

  for (const ext of extensions) {
    const runtime = ext.onSchemaReady(ctx);

    if (runtime.commands) {
      Object.assign(allCommands, runtime.commands);
    }

    if (runtime.keyboardShortcuts) {
      allKeyboardShortcuts.push(runtime.keyboardShortcuts);
    }

    if (runtime.plugins) {
      allPlugins.push(...runtime.plugins);
    }
  }

  // Build final plugin array:
  // Every editor runtime owns text input, and the caret a navigation key
  // moves, at the DOM/model boundary. Install
  // this centrally so body, header/footer, and note views share the policy.
  // Extension plugins and merged keymaps retain their relative priority.
  // The invariant check runs last so it sees what every other plugin's
  // appended transactions left behind. Off unless a test or development
  // build sets `globalThis.__folioTransactionInvariants`.
  const plugins = [
    createTextInputPlugin(),
    createNativeNavigationPlugin(),
    ...allPlugins,
    ...allKeyboardShortcuts.map((shortcuts) => keymap(shortcuts)),
    ...(areTransactionInvariantsEnabled()
      ? [createTransactionInvariantPlugin()]
      : []),
  ];

  return { plugins, commandRegistry: allCommands };
};
