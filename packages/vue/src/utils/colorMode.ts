/**
 * Color-mode resolution for the Vue host binding. Pure resolution stays
 * separate from Vue's provide/inject layer so browser capability checks and
 * system-preference subscriptions remain independently testable.
 *
 * @packageDocumentation
 */

/** UI color theme for the editor chrome and canvas. */
export type ColorMode = "light" | "dark" | "system";

const QUERY = "(prefers-color-scheme: dark)";

/** Whether `matchMedia` is usable in the current environment (false under SSR). */
function canMatchMedia(): boolean {
  return typeof window !== "undefined" && typeof window.matchMedia === "function";
}

/**
 * Current OS dark-mode preference. Returns false when `matchMedia` is
 * unavailable (e.g. server-side render), so it is safe to call as a state seed.
 *
 * @public
 */
export function prefersColorSchemeDark(): boolean {
  return canMatchMedia() ? window.matchMedia(QUERY).matches : false;
}

/**
 * Resolve the effective dark flag from a {@link ColorMode} and the current OS
 * preference. `'system'` follows the OS; `'dark'`/`'light'` are explicit.
 *
 * @public
 */
export function resolveIsDark(colorMode: ColorMode, systemDark: boolean): boolean {
  return colorMode === "dark" || (colorMode === "system" && systemDark);
}

/**
 * Subscribe to OS dark-mode changes. Invokes `onChange` immediately with the
 * current value (so a stale seed is corrected on entry), then on every change.
 * Returns an unsubscribe function. A no-op under SSR.
 *
 * @public
 */
export function subscribeSystemDark(onChange: (dark: boolean) => void): () => void {
  if (!canMatchMedia()) return () => {};
  const media = window.matchMedia(QUERY);
  onChange(media.matches);
  const listener = (e: MediaQueryListEvent) => onChange(e.matches);
  media.addEventListener("change", listener);
  return () => media.removeEventListener("change", listener);
}
