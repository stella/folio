import { Plugin } from "prosemirror-state";

import { clipboardTransformPlugin } from "../../internal/clipboardTransformPlugin";
import {
  createStarterKitExtensions,
  type StarterKitOptions,
} from "../../internal/starterKitExtensions";
import type { AnyExtension } from "./types";

export type { StarterKitOptions } from "../../internal/starterKitExtensions";

const starterKitClipboardPlugins = (props: Plugin["props"]) => {
  const plugins = clipboardTransformPlugin(props);
  const { handlePaste, handleDOMEvents } = props;
  const paste = handleDOMEvents?.paste;
  if (handlePaste !== undefined || paste !== undefined)
    plugins.push(
      new Plugin({
        props: {
          ...(handlePaste === undefined ? {} : { handlePaste }),
          ...(paste === undefined ? {} : { handleDOMEvents: { paste } }),
        },
      }),
    );
  return plugins;
};

/** Create the schema, commands, keymaps and clipboard rules for the DOCX editor. */
export function createStarterKit(
  options: StarterKitOptions = {},
): AnyExtension[] {
  const extensions = createStarterKitExtensions(options);
  for (const extension of extensions) {
    const { onSchemaReady } = extension;
    extension.onSchemaReady = (ctx) => {
      const { plugins, ...runtime } = onSchemaReady(ctx);
      if (plugins === undefined) return runtime;

      // Hosts that explicitly request a selection observer receive its view events.
      if (
        extension.config.name === "selectionTracker" &&
        options.onSelectionChange
      )
        return { ...runtime, plugins };

      return {
        ...runtime,
        plugins: plugins.flatMap(({ props }) =>
          starterKitClipboardPlugins(props),
        ),
      };
    };
  }
  return extensions;
}
