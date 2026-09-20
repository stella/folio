/**
 * GENERATED FILE — do not edit.
 *
 * The `ST_PresetLineDashVal` enumeration, in schema order, derived from
 * `specifications/generated/docx-transitional-schema.gen.json` by
 * `scripts/generate-preset-line-dash.ts`. Regenerate with:
 *
 *   bun run generate:preset-line-dash
 *
 * It is the dash vocabulary of `a:ln/a:prstDash@val`, and only that: the
 * spellings it shares with CSS `border-style` (`solid`, `dash`, `dot`)
 * name different patterns there. Read it through `./presetLineDash`.
 */

/** Every `ST_PresetLineDashVal` member, in schema order. */
export const PRESET_LINE_DASH_VALS = [
  "solid",
  "dot",
  "dash",
  "lgDash",
  "dashDot",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDot",
  "sysDashDot",
  "sysDashDotDot",
] as const;

/** `a:prstDash@val`: a preset dash the schema declares. */
export type PresetLineDashVal = (typeof PRESET_LINE_DASH_VALS)[number];
