import type { Plugin } from "prosemirror-state";

const UNSAFE_PLUGIN_KEY_PREFIXES = ["collab$", "history$", "y-sync$", "y-undo$"] as const;

const pluginRuntimeKey = (plugin: Plugin): string | null => {
  // ProseMirror exposes this key at runtime but marks it internal in its
  // declarations. Reading it reflectively lets this boundary recognize the
  // actual history and collaboration plugins, including suffixed instances.
  const key = Reflect.get(plugin, "key");
  return typeof key === "string" ? key : null;
};

const retainsOrTransportsSteps = (plugin: Plugin): boolean => {
  const key = pluginRuntimeKey(plugin);
  return key !== null && UNSAFE_PLUGIN_KEY_PREFIXES.some((prefix) => key.startsWith(prefix));
};

/** Build the private plugin set for synchronous headless review states. */
export const pluginsForHeadlessRevisionResolution = (
  plugins: readonly Plugin[],
): readonly Plugin[] => plugins.filter((plugin) => !retainsOrTransportsSteps(plugin));
