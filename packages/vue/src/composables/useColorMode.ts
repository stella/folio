import {
  computed,
  inject,
  provide,
  ref,
  toValue,
  watchEffect,
  type App,
  type ComputedRef,
  type InjectionKey,
  type MaybeRefOrGetter,
} from "vue";
import {
  prefersColorSchemeDark,
  resolveIsDark,
  subscribeSystemDark,
  type ColorMode,
} from "../utils/colorMode";

/** Color mode used when a host does not provide one. */
export const defaultColorMode: ColorMode = "light";

const COLOR_MODE_KEY: InjectionKey<ComputedRef<ColorMode>> = Symbol("folioColorMode");

/**
 * Provide a reactive color mode to descendant folio components.
 *
 * Call this from a parent component's setup function. For app-wide setup, use
 * {@link colorModePlugin} instead.
 */
export const provideColorMode = (
  colorMode: MaybeRefOrGetter<ColorMode> = defaultColorMode,
): void => {
  provide(
    COLOR_MODE_KEY,
    computed(() => toValue(colorMode)),
  );
};

/** Vue plugin form: `app.use(colorModePlugin, "system")`. */
export const colorModePlugin = {
  install(app: App, colorMode: MaybeRefOrGetter<ColorMode> = defaultColorMode): void {
    app.provide(
      COLOR_MODE_KEY,
      computed(() => toValue(colorMode)),
    );
  },
};

/**
 * Resolve the effective dark flag from the inherited color mode. `'system'`
 * follows the OS via `subscribeSystemDark` (SSR-safe; re-syncs on entry).
 */
export function useColorMode(): ComputedRef<boolean> {
  const colorMode = inject(
    COLOR_MODE_KEY,
    computed(() => defaultColorMode),
  );
  const systemDark = ref(prefersColorSchemeDark());
  watchEffect((onCleanup) => {
    if (colorMode.value !== "system") return;
    onCleanup(
      subscribeSystemDark((dark) => {
        systemDark.value = dark;
      }),
    );
  });
  return computed(() => resolveIsDark(colorMode.value, systemDark.value));
}

export type { ColorMode } from "../utils/colorMode";
