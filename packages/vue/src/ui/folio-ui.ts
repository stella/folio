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
 * Full parity with the React contract: all ten injectable primitives
 * (Button, Dialog, Select, Menu, Popover, Input, Checkbox, ColorPicker,
 * DatePickerPopover, OutlineRail) are present here with a real Vue default.
 *
 * Modeling note (vs. the React contract): React's Dialog/Select/Menu/Popover
 * are base-ui compound part-objects (`{ Root, Trigger, Popup, ... }`) a
 * consumer composes from sub-parts. The Vue adapter ships bespoke
 * single-file components instead of base-ui parts, so each primitive here is
 * one monolithic, data-driven Vue component (e.g. `Menu` takes a flat
 * `label` + `items` array rather than `Root`/`Trigger`/`Item` sub-parts) —
 * matching the shape the pre-existing Button/ColorPicker/Popover/Menu entries
 * already used. The exported `Folio*Props` types document each primitive's
 * prop contract; consumers override any subset through `DocxEditor`'s
 * `components` prop.
 *
 * Threading note: `DatePickerPopover`'s default exists and is injectable, but
 * no Vue chrome consumer renders it yet — the content-control widget overlay
 * (`ContentControlWidgetsOverlay` in React) that would use it is not ported to
 * Vue (a separate, pre-existing gap unrelated to UI injection).
 */

import { defineComponent, inject, provide, type Component, type InjectionKey } from "vue";

import DefaultButton from "../components/ui/Button.vue";
import DefaultCheckbox from "../components/ui/Checkbox.vue";
import DefaultColorPicker from "../components/ui/ColorPicker.vue";
import DefaultDatePickerPopover from "../components/ui/DatePickerPopover.vue";
import DefaultDialog from "../components/ui/Dialog.vue";
import DefaultInput from "../components/ui/Input.vue";
import DefaultMenuDropdown from "../components/ui/MenuDropdown.vue";
import DefaultOutlineRail from "../components/ui/OutlineRail.vue";
import DefaultPopover from "../components/ui/Popover.vue";
import DefaultSelect from "../components/ui/Select.vue";

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
// Dialog
// ============================================================================

/**
 * The Dialog prop subset folio's chrome relies on. Mirrors the Vue
 * `ui/Dialog.vue` surface: `open` controls the modal's visibility (the default
 * teleports to `document.body`, shows a click-to-close backdrop, and closes on
 * Escape); `ariaLabel` is applied to the popup's `role="dialog"` element.
 * `className` / `backdropClass` extend the popup/backdrop classes so a
 * dialog-specific stylesheet still applies over the shared default. Consumers
 * emit `update:open` / `close` (mirroring `Popover`'s `update:open` + `close`
 * pair) and render the dialog body through the default slot. A design-system
 * Dialog accepting a superset stays assignable as an override.
 */
export type FolioDialogProps = {
  open: boolean;
  ariaLabel?: string;
  className?: string;
  backdropClass?: string;
  /** Close on backdrop mousedown (default true). */
  closeOnBackdrop?: boolean;
};

// ============================================================================
// Select
// ============================================================================

/** One option in a folio Select. Mirrors the flat, data-driven shape the Vue
 *  contract already uses for {@link FolioMenuItem}. */
export type FolioSelectItem = {
  value: string;
  label: string;
  disabled?: boolean;
};

/**
 * The Select prop subset folio's chrome relies on. Mirrors the Vue
 * `ui/Select.vue` surface: a single native-`<select>`-backed component driven
 * by a flat `items` array (no listbox/positioner sub-parts, no `<optgroup>`
 * grouping — chrome consumers that need grouping, e.g. the font picker, stay
 * on a native `<select>` and are out of this contract's scope). `value` is the
 * selected item's `value`; the component emits `change` with the new value.
 */
export type FolioSelectProps = {
  value?: string;
  items: FolioSelectItem[];
  disabled?: boolean;
  placeholder?: string;
  className?: string;
};

// ============================================================================
// Input
// ============================================================================

/**
 * The Input prop subset folio's chrome relies on: a thin wrapper over a native
 * `<input>` supporting `v-model`. `className` / `size` are folio's shorthand
 * (mirroring React's `FolioInputProps`); every other native `<input>`
 * attribute (`type`, `placeholder`, `aria-*`, `min`/`max`/`step`, ...) falls
 * through Vue's attribute inheritance and is not re-declared.
 *
 * A chrome consumer that needs imperative access (e.g. FindReplaceDialog's
 * focus-and-select-on-open) does so through a template `ref`, which resolves
 * to the component instance — so an override, like the default, should
 * `defineExpose({ focus, select })` if it wants that behavior to keep working.
 */
export type FolioInputProps = {
  className?: string;
  size?: "sm" | "default" | "lg";
};

// ============================================================================
// Checkbox
// ============================================================================

/**
 * The Checkbox prop subset folio's chrome relies on. Mirrors React's
 * `FolioCheckboxProps` (`checked` + a change callback): `checked` is the
 * current state, and the component emits `update:checked` so a host uses
 * `v-model:checked`. `className` extends the default indicator's classes.
 */
export type FolioCheckboxProps = {
  checked?: boolean;
  className?: string;
};

// ============================================================================
// DatePickerPopover
// ============================================================================

/**
 * The DatePickerPopover prop subset folio's chrome relies on. Mirrors React's
 * `FolioDatePickerPopoverProps`: `value` accepts an ISO string, a `Date`, or
 * `null`; the component emits `change` with an ISO `yyyy-mm-dd` string (or
 * `null` when cleared). See the module docblock — no Vue chrome consumer
 * renders this yet (content-control widgets are not ported to Vue).
 */
export type FolioDatePickerPopoverProps = {
  value: string | Date | null;
  clearLabel?: string;
  defaultOpen?: boolean;
  showIcon?: boolean;
};

// ============================================================================
// OutlineRail
// ============================================================================

/**
 * The OutlineRail prop subset folio's chrome relies on. Mirrors React's
 * `FolioOutlineRailProps`, adapted to Vue's idiom for sharing a mutable DOM
 * ref across a prop boundary: a `getScrollContainer` getter (the same
 * `() => HTMLElement | null` shape `DecorationLayer.vue`'s
 * `getPagesContainer` already uses) rather than a raw `Ref` — Vue's template
 * compiler auto-unwraps a bare `Ref` referenced in a binding expression (even
 * inside an inline arrow function), so passing the ref object itself across a
 * prop is not the idiom here; a getter defers the read to click-time. The
 * rail resolves each item's vertical position via `resolvePct` and navigates
 * via `onJump` (both receive the resolved scroll container, as plain function
 * props — the same idiom the existing {@link FolioMenuItem}`.onSelect` uses).
 * `activeId` controls the highlighted entry. A design-system rail accepting a
 * superset stays assignable as an override.
 */
export type FolioOutlineRailProps = {
  items: OutlineItem[];
  getScrollContainer: () => HTMLElement | null;
  resolvePct?: (id: string, container: HTMLElement) => number | null;
  onJump: (id: string, container: HTMLElement) => void;
  activeId?: string | null;
  topOffset?: number;
  panelWidth?: number;
  ariaLabel?: string;
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
// OutlineRail — shared item type ({@link FolioOutlineRailProps} above)
// ============================================================================

/**
 * One entry in the document outline. Mirrors the React contract's item shape so
 * a cross-framework design system can share it.
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
 * exported `Folio*Props` types ({@link FolioButtonProps}, {@link FolioDialogProps},
 * {@link FolioSelectProps}, {@link FolioMenuProps}, {@link FolioPopoverProps},
 * {@link FolioInputProps}, {@link FolioCheckboxProps},
 * {@link FolioColorPickerProps}, {@link FolioDatePickerPopoverProps},
 * {@link FolioOutlineRailProps}); consumers render the resolved component and
 * pass those props at the call site.
 */
export type FolioUIComponents = {
  Button: Component;
  Dialog: Component;
  Select: Component;
  Menu: Component;
  Popover: Component;
  Input: Component;
  Checkbox: Component;
  ColorPicker: Component;
  DatePickerPopover: Component;
  OutlineRail: Component;
};

export const DEFAULT_COMPONENTS: FolioUIComponents = {
  Button: DefaultButton,
  Dialog: DefaultDialog,
  Select: DefaultSelect,
  Menu: DefaultMenuDropdown,
  Popover: DefaultPopover,
  Input: DefaultInput,
  Checkbox: DefaultCheckbox,
  ColorPicker: DefaultColorPicker,
  DatePickerPopover: DefaultDatePickerPopover,
  OutlineRail: DefaultOutlineRail,
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
