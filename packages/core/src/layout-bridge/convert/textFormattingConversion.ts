/**
 * Converts document-model TextFormatting (style, paragraph-mark, and list
 * marker run properties) to layout RunFormatting, including theme-font
 * resolution per script slot.
 */

import { getFontAlternate, type FontAlternates } from "../../fonts/fontAlternates";
import type { RunFormatting, ListMarkerFormatting } from "../../layout-engine/types";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import type { ParagraphAttrs as PMParagraphAttrs } from "../../prosemirror/schema/nodes";
import type { ColorValue, Theme, TextFormatting } from "../../types/document";
import { resolveColor, resolveHighlightToCss } from "../../utils/colorResolver";
import { resolveThemeFont } from "../../utils/fontResolver";
import { halfPointsToPixels, halfPointsToPoints } from "../../utils/units";
import { twipsToPixels } from "./flowConversionShared";

type ThemeFontAttributes = Omit<NonNullable<TextFormatting["fontFamily"]>, "asciiTheme"> & {
  /** The ProseMirror attr reader validates this against the canonical theme values. */
  asciiTheme?: string;
};

export const resolveWesternThemeFont = (
  fontFamily: ThemeFontAttributes,
  theme?: Theme | null,
): string | undefined => {
  const themeRef = fontFamily.asciiTheme ?? fontFamily.hAnsiTheme;
  const themedFont = themeRef ? resolveThemeFont(themeRef, theme?.fontScheme) : null;
  return themedFont ?? fontFamily.ascii ?? fontFamily.hAnsi ?? undefined;
};

export const resolveComplexScriptThemeFont = (
  fontFamily: ThemeFontAttributes,
  theme?: Theme | null,
): string | undefined => {
  // OOXML spells this attribute all-lowercase (`w:cstheme`), but the
  // canonical model uses the camelCase `csTheme` field.
  const themedFont = fontFamily.csTheme
    ? resolveThemeFont(fontFamily.csTheme, theme?.fontScheme)
    : null;
  return themedFont ?? fontFamily.cs ?? undefined;
};

export const resolveEastAsiaThemeFont = (
  fontFamily: ThemeFontAttributes,
  theme?: Theme | null,
): string | undefined => {
  const themedFont = fontFamily.eastAsiaTheme
    ? resolveThemeFont(fontFamily.eastAsiaTheme, theme?.fontScheme)
    : null;
  return themedFont ?? fontFamily.eastAsia ?? undefined;
};

export function isAutomaticTextColorValue(color: ColorValue): boolean {
  const rgb = color.rgb?.trim().toLowerCase();
  return color.auto === true || rgb === "auto" || (!rgb && !color.themeColor);
}

function textFormattingToRunFormatting(
  defaultTextFormatting: TextFormatting | undefined,
  theme?: Theme | null,
  fontAlternates?: FontAlternates,
): RunFormatting {
  if (!defaultTextFormatting) {
    return {};
  }

  const result: RunFormatting = {};
  const fontFamily = defaultTextFormatting.fontFamily
    ? resolveWesternThemeFont(defaultTextFormatting.fontFamily, theme)
    : undefined;
  if (fontFamily) {
    result.fontFamily = fontFamily;
    const alternate = getFontAlternate(fontFamily, fontAlternates);
    if (alternate) {
      result.alternateFontFamily = alternate;
    }
  }
  // East-Asian font inherited from the paragraph style / docDefaults, so CJK
  // runs without a direct `w:eastAsia` still get per-character EA selection. A
  // run's own fontFamily mark overrides this via mergeRunFormatting.
  const eastAsiaFontFamily = defaultTextFormatting.fontFamily
    ? resolveEastAsiaThemeFont(defaultTextFormatting.fontFamily, theme)
    : undefined;
  if (eastAsiaFontFamily) {
    result.eastAsiaFontFamily = eastAsiaFontFamily;
    const alternate = getFontAlternate(eastAsiaFontFamily, fontAlternates);
    if (alternate) {
      result.eastAsiaAlternateFontFamily = alternate;
    }
  }
  if (defaultTextFormatting.fontFamily?.hint) {
    result.eastAsiaHint = defaultTextFormatting.fontFamily.hint === "eastAsia";
  }
  const complexScriptFontFamily = defaultTextFormatting.fontFamily
    ? resolveComplexScriptThemeFont(defaultTextFormatting.fontFamily, theme)
    : undefined;
  if (complexScriptFontFamily) {
    result.complexScriptFontFamily = complexScriptFontFamily;
    const alternate = getFontAlternate(complexScriptFontFamily, fontAlternates);
    if (alternate) {
      result.complexScriptAlternateFontFamily = alternate;
    }
  }
  if (defaultTextFormatting.language) {
    result.language = { ...defaultTextFormatting.language };
  }
  if (defaultTextFormatting.fontSize !== undefined) {
    result.fontSize = defaultTextFormatting.fontSize / 2;
  }
  if (defaultTextFormatting.fontSizeCs !== undefined) {
    result.complexScriptFontSize = defaultTextFormatting.fontSizeCs / 2;
  }
  if (defaultTextFormatting.bold !== undefined) {
    result.bold = defaultTextFormatting.bold;
  }
  if (defaultTextFormatting.boldCs !== undefined) {
    result.complexScriptBold = defaultTextFormatting.boldCs;
  }
  if (defaultTextFormatting.italic !== undefined) {
    result.italic = defaultTextFormatting.italic;
  }
  if (defaultTextFormatting.italicCs !== undefined) {
    result.complexScriptItalic = defaultTextFormatting.italicCs;
  }
  if (defaultTextFormatting.rtl !== undefined) {
    result.rtl = defaultTextFormatting.rtl;
  }
  if (defaultTextFormatting.cs !== undefined) {
    result.forceComplexScript = defaultTextFormatting.cs;
  }
  if (defaultTextFormatting.underline && defaultTextFormatting.underline.style !== "none") {
    result.underline = { style: defaultTextFormatting.underline.style };
    if (defaultTextFormatting.underline.color) {
      result.underline.color = resolveColor(defaultTextFormatting.underline.color, theme);
    }
  }
  if (defaultTextFormatting.strike !== undefined) {
    result.strike = defaultTextFormatting.strike;
  }
  if (defaultTextFormatting.color && !isAutomaticTextColorValue(defaultTextFormatting.color)) {
    result.color = resolveColor(defaultTextFormatting.color, theme);
    result.textColorSource = "paragraphDefault";
  }
  if (defaultTextFormatting.highlight) {
    const highlight = resolveHighlightToCss(defaultTextFormatting.highlight);
    if (highlight) {
      result.highlight = highlight;
    }
  }
  if (defaultTextFormatting.vertAlign === "superscript") {
    result.superscript = true;
  }
  if (defaultTextFormatting.vertAlign === "subscript") {
    result.subscript = true;
  }
  if (defaultTextFormatting.allCaps !== undefined) {
    result.allCaps = defaultTextFormatting.allCaps;
  }
  if (defaultTextFormatting.smallCaps !== undefined) {
    result.smallCaps = defaultTextFormatting.smallCaps;
  }
  if (defaultTextFormatting.spacing !== undefined && defaultTextFormatting.spacing !== 0) {
    result.letterSpacing = twipsToPixels(defaultTextFormatting.spacing);
  }
  if (defaultTextFormatting.position !== undefined && defaultTextFormatting.position !== 0) {
    result.positionPx = halfPointsToPixels(defaultTextFormatting.position);
  }
  const horizontalScale = normalizeHorizontalScalePercent(defaultTextFormatting.scale);
  if (horizontalScale !== undefined && horizontalScale !== 100) {
    result.horizontalScale = horizontalScale;
  }
  if (defaultTextFormatting.kerning !== undefined && defaultTextFormatting.kerning > 0) {
    result.kerningMinPt = halfPointsToPoints(defaultTextFormatting.kerning);
  }
  if (defaultTextFormatting.emboss !== undefined) {
    result.emboss = defaultTextFormatting.emboss;
  }
  if (defaultTextFormatting.imprint !== undefined) {
    result.imprint = defaultTextFormatting.imprint;
  }
  if (defaultTextFormatting.shadow !== undefined) {
    result.textShadow = defaultTextFormatting.shadow;
  }
  if (defaultTextFormatting.outline !== undefined) {
    result.textOutline = defaultTextFormatting.outline;
  }
  if (defaultTextFormatting.emphasisMark && defaultTextFormatting.emphasisMark !== "none") {
    result.emphasisMark = defaultTextFormatting.emphasisMark;
  }
  return result;
}

const LIST_MARKER_FORMATTING_KEYS = [
  "fontFamily",
  "alternateFontFamily",
  "eastAsiaFontFamily",
  "eastAsiaAlternateFontFamily",
  "complexScriptFontFamily",
  "complexScriptAlternateFontFamily",
  "fontSize",
  "complexScriptFontSize",
  "bold",
  "complexScriptBold",
  "italic",
  "complexScriptItalic",
  "rtl",
  "forceComplexScript",
  "color",
] as const satisfies readonly (keyof ListMarkerFormatting)[];

/**
 * The numbering symbol's typography: the numbering level's `w:rPr` applied
 * over the paragraph mark's run properties (ECMA-376 §17.9.24, §17.3.1.29).
 * Text runs of the paragraph play no part.
 */
export function listMarkerFormattingFor(
  levelFormatting: TextFormatting | undefined,
  paragraphMarkFormatting: TextFormatting | undefined,
  theme: Theme | null | undefined,
  fontAlternates: FontAlternates | undefined,
): ListMarkerFormatting | undefined {
  const runFormatting = textFormattingToRunFormatting(
    mergeTextFormatting(paragraphMarkFormatting, levelFormatting),
    theme,
    fontAlternates,
  );
  const formatting: ListMarkerFormatting = {};
  for (const key of LIST_MARKER_FORMATTING_KEYS) {
    if (runFormatting[key] !== undefined) {
      Reflect.set(formatting, key, runFormatting[key]);
    }
  }
  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

export function paragraphRunDefaults(
  pmAttrs: PMParagraphAttrs,
  theme?: Theme | null,
  fontAlternates?: FontAlternates,
): RunFormatting {
  return textFormattingToRunFormatting(pmAttrs.defaultTextFormatting, theme, fontAlternates);
}
