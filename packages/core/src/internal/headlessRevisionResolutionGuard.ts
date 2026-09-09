import { Plugin, PluginKey, type EditorState } from "prosemirror-state";

const UNSAFE_PLUGIN_KEY_PREFIXES = ["collab$", "history$", "y-sync$", "y-undo$"] as const;

const headlessRevisionResolutionKey = new PluginKey("folioHeadlessRevisionResolution");
const headlessRevisionResolutionPlugin = new Plugin({ key: headlessRevisionResolutionKey });

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

/**
 * Whether a state can consume an unmapped replacement before it returns.
 * Only the internal headless plugin set opts in; history and collaboration
 * plugins always keep the state on the legacy editor-command path.
 */
export const stateAllowsHeadlessRevisionResolution = (state: EditorState): boolean =>
  headlessRevisionResolutionKey.get(state) === headlessRevisionResolutionPlugin &&
  !state.plugins.some(retainsOrTransportsSteps);

/** Build the private plugin set for synchronous headless review states. */
export const pluginsForHeadlessRevisionResolution = (
  plugins: readonly Plugin[],
): readonly Plugin[] => [
  ...plugins.filter((plugin) => !retainsOrTransportsSteps(plugin)),
  headlessRevisionResolutionPlugin,
];
