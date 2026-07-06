/**
 * Types for `AlignmentButtons.vue`, kept in a plain `.ts` module so downstream
 * `.ts` consumers (and the standalone `FormattingBar`) can import the type
 * without resolving a `.vue` SFC. A plain `tsc` (e.g. the Nuxt adapter's
 * typecheck) sees `.vue` imports only through an ambient default-export shim, so
 * named type exports must live in `.ts` files.
 */

/**
 * Paragraph alignment the picker offers. Intentionally the narrow subset the
 * minimal toolbar exposes; core's `ParagraphAlignment` is a wider superset
 * (distribute + Kashida variants) that this control does not surface.
 */
export type ParagraphAlignment = "left" | "center" | "right" | "both";
