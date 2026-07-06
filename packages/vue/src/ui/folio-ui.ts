/**
 * Vue UI-injection layer (mirrors `packages/react/src/ui/folio-ui.tsx`).
 *
 * folio's chrome (toolbars, pickers, dialogs) renders a small set of UI
 * primitives a host can override with its own design system. React models this
 * as a {@link FolioUIComponents} record injected through a context provider and
 * read with `useFolioUI`; the Vue adapter mirrors the contract with Vue's
 * provide/inject.
 *
 * Standalone folio uses {@link DEFAULT_COMPONENTS} (the package's own `ui/*`
 * single-file components); consumers inject overrides through `DocxEditor`'s
 * `components` prop, which calls {@link provideFolioUI}.
 *
 * Scope note (vs. the React contract): the React map ships ten primitives, some
 * as base-ui compound part-objects (Dialog/Select/Menu/Popover). The Vue adapter
 * ships bespoke single-file components rather than base-ui parts, so the Vue
 * contract carries the primitives that have a real Vue default and are consumed
 * by the ported chrome: Button, ColorPicker, Popover, and Menu. The remaining
 * React primitives (Dialog, Select, Input, Checkbox, DatePickerPopover,
 * OutlineRail) have no Vue default primitive to inject yet; their shared value
 * types ({@link ColorPreset}, {@link OutlineItem}) are still exported for
 * cross-framework parity.
 */

import { defineComponent, inject, provide, type Component, type InjectionKey } from "vue";

import DefaultButton from "../components/ui/Button.vue";
import DefaultColorPicker from "../components/ui/ColorPicker.vue";
import DefaultMenuDropdown from "../components/ui/MenuDropdown.vue";
import DefaultPopover from "../components/ui/Popover.vue";

// ============================================================================
// Button
// ============================================================================

/**
 * The Button prop subset folio's chrome relies on. `variant` and `size` are
 * narrow string-literal unions (a subset of the design-system Button's options)
 * so an external Button stays assignable as an override. Native `<button>`
 * attributes fall through Vue's attribute inheritance and are not re-declared.
 */
export type FolioButtonProps = {
  variant?: "default" | "ghost";
  size?: "sm" | "xs" | "icon-xs";
  disabled?: boolean;
  className?: string;
};

// ============================================================================
// ColorPicker
// ============================================================================

/**
 * A color preset rendered as a swatch in the picker. Mirrors the React contract
 * so a cross-framework design system can share the preset shape; `color` is the
 * optional CSS color for the swatch (falls back to `#${value}`), and `value` is
 * what selection emits.
 */
export type ColorPreset = {
  label: string;
  value: string;
  color?: string;
};

/**
 * The ColorPicker prop subset folio's chrome relies on. Mirrors the Vue
 * `ui/ColorPicker.vue` surface: `mode` selects the text/highlight/border
 * behaviour, `value` is the current color, `theme` resolves theme colors, and
 * the component emits `change` with the picked color. A design-system picker
 * accepting a superset stays assignable as an override.
 */
export type FolioColorPickerProps = {
  mode: "text" | "highlight" | "border";
  value?: string | undefined;
  disabled?: boolean;
  className?: string;
};

// ============================================================================
// Popover
// ============================================================================

/**
 * The Popover prop subset folio's chrome relies on. Mirrors the Vue
 * `ui/Popover.vue` surface: `open` controls visibility, `placement` positions
 * the panel, and the component emits `close` / `update:open`. Consumers supply
 * the trigger and panel via the `trigger` / `panel` slots.
 */
export type FolioPopoverProps = {
  open: boolean;
  placement?: "bottom-left" | "bottom-right" | "top-left" | "top-right";
  closeOnScroll?: boolean;
};

// ============================================================================
// Menu
// ============================================================================

/**
 * One entry in a folio menu. Mirrors the Vue `ui/MenuDropdown.vue` `MenuEntry`
 * shape so an external Menu stays assignable as an override.
 */
export type FolioMenuItem = {
  type?: "item";
  label: string;
  onSelect?: () => void;
  disabled?: boolean;
  shortcut?: string;
};

export type FolioMenuProps = {
  label: string;
  items: ReadonlyArray<FolioMenuItem | { type: "separator" }>;
};

// ============================================================================
// OutlineRail (shared value type only; no Vue default primitive yet)
// ============================================================================

/**
 * One entry in the document outline. Mirrors the React contract's item shape so
 * a cross-framework design system can share it. The Vue adapter does not yet
 * inject an OutlineRail primitive (no Vue default exists); the type is exported
 * for parity with the React surface.
 */
export type OutlineItem = {
  id: string;
  label: string;
  /** Nesting depth; drives indent + tick taper. */
  level: number;
  /** Optional trailing annotation in the panel (e.g. a page number). */
  meta?: string;
  /** Optional CSS custom-property name colouring this entry. */
  color?: string;
};

// ============================================================================
// Contract + provide/inject
// ============================================================================

/**
 * The injectable chrome primitives. Each entry is a Vue component the chrome
 * renders; a host overrides any subset through `DocxEditor`'s `components` prop.
 *
 * Values are typed as the loose Vue {@link Component} rather than
 * `Component<FolioButtonProps>` etc.: Vue's `Component<P>` is invariant in `P`
 * (unlike React's contravariant `ComponentType<P>`), so a default or override
 * accepting a *superset* of props would not be assignable to a prop-parameterized
 * slot. The per-primitive prop contract a host implements is documented by the
 * exported `Folio*Props` types ({@link FolioButtonProps},
 * {@link FolioColorPickerProps}, {@link FolioPopoverProps}, {@link FolioMenuProps});
 * consumers render the resolved component and pass those props at the call site.
 */
export type FolioUIComponents = {
  Button: Component;
  ColorPicker: Component;
  Popover: Component;
  Menu: Component;
};

export const DEFAULT_COMPONENTS: FolioUIComponents = {
  Button: DefaultButton,
  ColorPicker: DefaultColorPicker,
  Popover: DefaultPopover,
  Menu: DefaultMenuDropdown,
};

const FOLIO_UI_KEY: InjectionKey<FolioUIComponents> = Symbol("folio-ui-components");

/**
 * Merge a consumer's partial overrides over the built-in defaults: every key is
 * present (defaults fill the gaps), and each provided override wins. Pure, so
 * the resolution contract is testable independently of Vue.
 */
export function resolveFolioComponents(components?: Partial<FolioUIComponents>): FolioUIComponents {
  return components ? { ...DEFAULT_COMPONENTS, ...components } : DEFAULT_COMPONENTS;
}

/**
 * Provide the resolved primitive map to descendants. `DocxEditor` calls this in
 * setup with its `components` prop; chrome descendants read it via
 * {@link useFolioUI}.
 */
export function provideFolioUI(components?: Partial<FolioUIComponents>): void {
  provide(FOLIO_UI_KEY, resolveFolioComponents(components));
}

/**
 * Read the resolved primitive map. Falls back to {@link DEFAULT_COMPONENTS} when
 * no provider is present (chrome rendered standalone, outside `DocxEditor`).
 */
export function useFolioUI(): FolioUIComponents {
  return inject(FOLIO_UI_KEY, DEFAULT_COMPONENTS);
}

/**
 * Provider component mirroring React's `FolioUIProvider`. Wraps a subtree and
 * provides the resolved primitive map; renders its default slot unchanged.
 * Defined here (a `.ts` module) rather than as a `.vue` SFC so it re-exports as
 * a named export the Nuxt `tsc` typecheck can resolve.
 */
export const FolioUIProvider = defineComponent({
  name: "FolioUIProvider",
  props: {
    components: {
      type: Object as () => Partial<FolioUIComponents> | undefined,
      default: undefined,
    },
  },
  setup(props, { slots }) {
    provideFolioUI(props.components);
    return () => slots["default"]?.();
  },
});
