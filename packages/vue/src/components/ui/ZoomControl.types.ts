/**
 * Types for `ZoomControl.vue`, kept in a plain `.ts` module so the package's
 * public `index.ts` can re-export them without resolving a `.vue` SFC. A plain
 * `tsc` (e.g. the Nuxt adapter's typecheck) sees `.vue` imports only through an
 * ambient default-export shim, so named type exports must live in `.ts` files.
 */

/** A single zoom preset: `value` is a scale factor (1 = 100%). */
export type ZoomLevel = {
  value: number;
  label: string;
};

export type ZoomControlProps = {
  /** Current zoom (1 = 100%). */
  value?: number;
  /** Override the preset levels offered in the dropdown. */
  levels?: ZoomLevel[];
  disabled?: boolean;
  className?: string;
  /** Render the trigger at the smaller toolbar-chrome size. */
  compact?: boolean;
};
