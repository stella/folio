import { panic } from "better-result";
import { PluginKey } from "prosemirror-state";
import type { EditorState, Plugin } from "prosemirror-state";

const PLUGIN_KEY_PREFIX = "stll.folio.plugin.";
const METADATA_KEY_PREFIX = "stll.folio.meta.";

export type PluginKeyOwner<State> = Readonly<{
  key: string;
  createKey: () => PluginKey<State>;
  owns: (plugin: Plugin<State>) => boolean;
  get: (state: EditorState) => Plugin<State> | undefined;
  getState: (state: EditorState) => State | undefined;
}>;

/** An immutable plugin family descriptor; each plugin owns its actual key. */
export const createPluginKeyOwner = <State = unknown>(name: string): PluginKeyOwner<State> => {
  const keys = new WeakSet<PluginKey<State>>();
  const pluginsByState = new WeakMap<EditorState, Plugin<State> | undefined>();
  const owns = (plugin: Plugin<State>): boolean =>
    plugin.spec.key !== undefined && keys.has(plugin.spec.key);
  const get = (state: EditorState): Plugin<State> | undefined => {
    if (pluginsByState.has(state)) return pluginsByState.get(state);
    let found: Plugin<State> | undefined;
    for (const plugin of state.plugins) {
      if (!owns(plugin)) continue;
      if (found) panic(`Multiple plugins belong to ${name} in one editor state`);
      found = plugin;
    }
    pluginsByState.set(state, found);
    return found;
  };
  return Object.freeze({
    key: `${METADATA_KEY_PREFIX}${name}`,
    createKey: (): PluginKey<State> => {
      const key = new PluginKey<State>(`${PLUGIN_KEY_PREFIX}${name}`);
      keys.add(key);
      return key;
    },
    owns,
    get,
    getState: (state: EditorState): State | undefined => get(state)?.getState(state),
  });
};
