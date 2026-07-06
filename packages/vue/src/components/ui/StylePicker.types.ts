/**
 * Types for `StylePicker.vue`, kept in a plain `.ts` module so downstream `.ts`
 * consumers (and the standalone `FormattingBar`) can import the type without
 * resolving a `.vue` SFC. A plain `tsc` (e.g. the Nuxt adapter's typecheck) sees
 * `.vue` imports only through an ambient default-export shim, so named type
 * exports must live in `.ts` files.
 */

/** A single paragraph-style option the picker can offer. */
export type StyleOption = {
  styleId: string;
  name: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
};
