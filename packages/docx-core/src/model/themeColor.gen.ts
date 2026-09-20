/**
 * GENERATED FILE — do not edit.
 *
 * The colour enumerations a theme reference passes through, derived from
 * `specifications/generated/docx-transitional-schema.gen.json` by
 * `scripts/generate-theme-colors.ts`. Regenerate with:
 *
 *   bun run generate:theme-colors
 */

/**
 * `ST_ThemeColor`: every token a WordprocessingML `w:themeColor`,
 * `w:themeFill` or `w:clrSchemeMapping` value may carry.
 *
 * Spelled in full words (`dark1`, `hyperlink`), unlike the DrawingML slot
 * names below. `none` is a reserved member that cancels an inherited theme
 * colour rather than naming one.
 */
export const THEME_COLORS = [
  "dark1",
  "light1",
  "dark2",
  "light2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hyperlink",
  "followedHyperlink",
  "none",
  "background1",
  "text1",
  "background2",
  "text2",
] as const;

export type ThemeColor =
  | "dark1"
  | "light1"
  | "dark2"
  | "light2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink"
  | "none"
  | "background1"
  | "text1"
  | "background2"
  | "text2";

/**
 * The colour slots a theme part's `a:clrScheme` declares, in schema order.
 *
 * This is the set a theme lookup can hit; {@link SCHEME_COLOR_VALUES} is wider.
 */
export const SCHEME_COLOR_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export type SchemeColorSlot =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

/**
 * `ST_SchemeColorVal`: every token an `a:schemeClr/@val` reference may carry.
 *
 * A superset of {@link SCHEME_COLOR_SLOTS}: it adds the four mapped spellings
 * (`bg1`, `tx1`, `bg2`, `tx2`) and `phClr`, the placeholder a style
 * definition resolves against its instantiating context.
 */
export const SCHEME_COLOR_VALUES = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
  "phClr",
  "dk1",
  "lt1",
  "dk2",
  "lt2",
] as const;

export type SchemeColorValue =
  | "bg1"
  | "tx1"
  | "bg2"
  | "tx2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink"
  | "phClr"
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2";

/**
 * The attributes `w:clrSchemeMapping` carries in `settings.xml`, each naming
 * the theme slot one mapped colour resolves to.
 */
export const CLR_SCHEME_MAPPING_KEYS = [
  "bg1",
  "t1",
  "bg2",
  "t2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hyperlink",
  "followedHyperlink",
] as const;

export type ClrSchemeMappingKey =
  | "bg1"
  | "t1"
  | "bg2"
  | "t2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink";
