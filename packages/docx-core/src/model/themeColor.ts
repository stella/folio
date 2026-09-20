/**
 * The one mapping between a WordprocessingML theme reference and a theme slot.
 *
 * `w:themeColor` names a colour in `ST_ThemeColor`'s spelling; a theme part
 * declares its slots in DrawingML's. Three kinds of member, per ECMA-376
 * §17.18.97:
 *
 *   direct    `dark1`, `hyperlink`, `accent3`  name a slot outright
 *   mapped    `background1`, `text1`, …        name a `w:clrSchemeMapping` key,
 *                                              which `settings.xml` points at a
 *                                              slot (§17.15.1.20)
 *   cancel    `none`                           states that no theme colour
 *                                              applies, overriding an inherited
 *                                              one
 *
 * Every consumer reads the same map, so the DrawingML spelling appears once.
 */

import {
  type ClrSchemeMappingKey,
  SCHEME_COLOR_VALUES,
  type SchemeColorSlot,
  type SchemeColorValue,
  type ThemeColor,
  THEME_COLORS,
} from "./themeColor.gen";

/** What a theme reference resolves against, before a theme is consulted. */
export type ThemeColorTarget =
  | { readonly kind: "slot"; readonly slot: SchemeColorSlot }
  | { readonly kind: "mapped"; readonly mapping: ClrSchemeMappingKey }
  | { readonly kind: "cancel" };

const slot = (value: SchemeColorSlot): ThemeColorTarget => ({ kind: "slot", slot: value });
const mapped = (value: ClrSchemeMappingKey): ThemeColorTarget => ({
  kind: "mapped",
  mapping: value,
});

/**
 * Total over `ST_ThemeColor`: a token added by a schema refresh cannot land
 * without a decision about which slot it paints.
 */
export const THEME_COLOR_TARGETS = {
  dark1: slot("dk1"),
  light1: slot("lt1"),
  dark2: slot("dk2"),
  light2: slot("lt2"),
  accent1: slot("accent1"),
  accent2: slot("accent2"),
  accent3: slot("accent3"),
  accent4: slot("accent4"),
  accent5: slot("accent5"),
  accent6: slot("accent6"),
  hyperlink: slot("hlink"),
  followedHyperlink: slot("folHlink"),
  none: { kind: "cancel" },
  background1: mapped("bg1"),
  text1: mapped("t1"),
  background2: mapped("bg2"),
  text2: mapped("t2"),
} as const satisfies Record<ThemeColor, ThemeColorTarget>;

/**
 * The mapping a document carries when `settings.xml` declares none.
 *
 * Word writes `w:clrSchemeMapping` with exactly these pairs into every new
 * document, so it is the mapping a reader assumes rather than a folio default.
 */
export const DEFAULT_CLR_SCHEME_MAPPING = {
  bg1: "lt1",
  t1: "dk1",
  bg2: "lt2",
  t2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
} as const satisfies Record<ClrSchemeMappingKey, SchemeColorSlot>;

/**
 * The `a:schemeClr/@val` spelling of each theme colour.
 *
 * DrawingML keeps the mapped colours as their own tokens (`bg1`, `tx1`) rather
 * than resolving them, so the two vocabularies differ only in spelling. `none`
 * has no DrawingML equivalent: a scheme reference that paints nothing is
 * written by omitting the fill, not by naming a colour.
 */
export const SCHEME_COLOR_VALUE_BY_THEME_COLOR = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
  none: null,
  background1: "bg1",
  text1: "tx1",
  background2: "bg2",
  text2: "tx2",
} as const satisfies Record<ThemeColor, SchemeColorValue | null>;

/**
 * The inverse: how an `a:schemeClr/@val` reads as a WordprocessingML theme
 * colour. `phClr` is the placeholder a style definition resolves against its
 * instantiating context, which no `w:themeColor` can name.
 */
export const THEME_COLOR_BY_SCHEME_COLOR_VALUE = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hyperlink",
  folHlink: "followedHyperlink",
  phClr: null,
  dk1: "dark1",
  lt1: "light1",
  dk2: "dark2",
  lt2: "light2",
} as const satisfies Record<SchemeColorValue, ThemeColor | null>;

/**
 * A `w:themeColor` token outside `ST_ThemeColor`, kept as written.
 *
 * Narrowing used to drop such a token, which also dropped the attribute the
 * next time the element was written. Capturing it keeps the save leg lossless
 * while leaving the painter free to ignore a colour it cannot resolve — the
 * treatment `w:val` on `CT_Border` already gets for a style outside
 * `ST_Border`.
 */
export type UnrecognisedThemeColor = { readonly kind: "unrecognised"; readonly raw: string };

/** What a theme-colour attribute parses to: a schema token, or the raw text. */
export type ThemeColorValue = ThemeColor | UnrecognisedThemeColor;

const THEME_COLOR_SET: ReadonlySet<string> = new Set(THEME_COLORS);

export const isThemeColor = (value: string): value is ThemeColor => THEME_COLOR_SET.has(value);

const SCHEME_COLOR_VALUE_SET: ReadonlySet<string> = new Set(SCHEME_COLOR_VALUES);

export const isSchemeColorValue = (value: string): value is SchemeColorValue =>
  SCHEME_COLOR_VALUE_SET.has(value);

/**
 * Read an attribute as a theme colour, or `undefined` when it is absent.
 *
 * An unrecognised token is captured rather than refused; the caller decides
 * whether to report it.
 */
export const readThemeColor = (raw: string | null | undefined): ThemeColorValue | undefined => {
  if (raw === null || raw === undefined || raw === "") {
    return undefined;
  }
  return isThemeColor(raw) ? raw : { kind: "unrecognised", raw };
};

/** The token to write back, whether or not the schema enumerates it. */
export const themeColorToken = (value: ThemeColorValue): string =>
  typeof value === "string" ? value : value.raw;

/** The schema token, or `undefined` for one the model only carries verbatim. */
export const knownThemeColor = (value: ThemeColorValue): ThemeColor | undefined =>
  typeof value === "string" ? value : undefined;

/**
 * The theme slot a reference paints with, or `undefined` when it names none.
 *
 * `undefined` covers both `none`, which cancels an inherited theme colour, and
 * a token the schema does not enumerate.
 */
export const themeColorSlot = (
  value: ThemeColorValue,
  mapping: Readonly<Record<ClrSchemeMappingKey, SchemeColorSlot>> = DEFAULT_CLR_SCHEME_MAPPING,
): SchemeColorSlot | undefined => {
  const known = knownThemeColor(value);
  if (known === undefined) {
    return undefined;
  }
  const target: ThemeColorTarget = THEME_COLOR_TARGETS[known];
  switch (target.kind) {
    case "slot":
      return target.slot;
    case "mapped":
      return mapping[target.mapping];
    case "cancel":
      return undefined;
    default: {
      const exhaustive: never = target;
      return exhaustive;
    }
  }
};
