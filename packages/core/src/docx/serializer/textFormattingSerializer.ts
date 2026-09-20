import type {
  ColorValue,
  ExhaustiveFields,
  ShadingProperties,
  TextFormatting,
} from "../../types/document";
import { HIGHLIGHT_COLOR_VALUES } from "../../types/documentEnumValues";
import { isValidHexColor } from "../../utils/colorResolver";
import { roundHorizontalScalePercentForSerialization } from "../../utils/horizontalScale";
import type { DeclaredChild } from "../containerChildren.gen";
import { serializeSequenceChildren } from "../containerChildren";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute } from "@stll/docx-core";

const VALID_HIGHLIGHT_COLORS = new Set(HIGHLIGHT_COLOR_VALUES);

type ClassifiedColorField = "rgb" | "themeColor" | "themeTint" | "themeShade" | "auto";
type ExhaustiveColorValue = ExhaustiveFields<ColorValue, ClassifiedColorField>;

type ClassifiedShadingField = "color" | "fill" | "pattern";
type ExhaustiveShadingProperties = ExhaustiveFields<ShadingProperties, ClassifiedShadingField>;

type FontFamily = NonNullable<TextFormatting["fontFamily"]>;
type ClassifiedFontFamilyField =
  | "ascii"
  | "hAnsi"
  | "eastAsia"
  | "cs"
  | "hint"
  | "asciiTheme"
  | "hAnsiTheme"
  | "eastAsiaTheme"
  | "csTheme";
type ExhaustiveFontFamily = ExhaustiveFields<FontFamily, ClassifiedFontFamilyField>;

type Underline = NonNullable<TextFormatting["underline"]>;
type ClassifiedUnderlineField = "style" | "color";
type ExhaustiveUnderline = ExhaustiveFields<Underline, ClassifiedUnderlineField>;

type Language = NonNullable<TextFormatting["language"]>;
type ClassifiedLanguageField = "val" | "eastAsia" | "bidi";
type ExhaustiveLanguage = ExhaustiveFields<Language, ClassifiedLanguageField>;

type ClassifiedTextFormattingField =
  | "bold"
  | "boldCs"
  | "italic"
  | "italicCs"
  | "underline"
  | "strike"
  | "doubleStrike"
  | "vertAlign"
  | "smallCaps"
  | "allCaps"
  | "hidden"
  | "noProof"
  | "color"
  | "highlight"
  | "shading"
  | "fontSize"
  | "fontSizeCs"
  | "fontFamily"
  | "language"
  | "spacing"
  | "position"
  | "scale"
  | "kerning"
  | "effect"
  | "emphasisMark"
  | "emboss"
  | "imprint"
  | "outline"
  | "shadow"
  | "rtl"
  | "cs"
  | "styleId"
  | "preserved";
type ExhaustiveTextFormatting = ExhaustiveFields<TextFormatting, ClassifiedTextFormattingField>;

/**
 * A child of this `w:rPr` that another record owns, named by its declared
 * child so the sequence places it.
 *
 * A run's `w:rPrChange` and the paragraph mark's revision and `w:specVanish`
 * are read by the run's and the paragraph's own records, so they arrive here
 * as markup rather than as a field. Naming each one means the schema decides
 * where it goes, exactly as it does for the modelled children.
 */
type OwnedRunPropertyChild = readonly [name: DeclaredChild<"run-properties">, xml: string];

// ============================================================================
// COLOR SERIALIZATION
// ============================================================================

/**
 * Serialize a color element (w:color)
 */
function serializeColorElement(color: ExhaustiveColorValue | undefined): string {
  if (!color) {
    return "";
  }

  const { auto, rgb, themeColor, themeTint, themeShade } = color;

  const attrs: string[] = [];

  if (auto) {
    attrs.push('w:val="auto"');
  } else if (rgb && isValidHexColor(rgb)) {
    attrs.push(`w:val="${escapeXmlAttribute(rgb)}"`);
  }

  if (themeColor) {
    attrs.push(`w:themeColor="${escapeXmlAttribute(themeColor)}"`);
  }

  if (themeTint) {
    attrs.push(`w:themeTint="${escapeXmlAttribute(themeTint)}"`);
  }

  if (themeShade) {
    attrs.push(`w:themeShade="${escapeXmlAttribute(themeShade)}"`);
  }

  return attrs.length === 0 ? "" : `<w:color ${attrs.join(" ")}/>`;
}

// ============================================================================
// SHADING SERIALIZATION
// ============================================================================

/**
 * Serialize shading properties (w:shd)
 */
export function serializeShading(shading: ExhaustiveShadingProperties | undefined): string {
  if (!shading) {
    return "";
  }

  const { pattern, color, fill } = shading;

  const attrs: string[] = [];

  // Pattern/val
  if (pattern) {
    attrs.push(`w:val="${escapeXmlAttribute(pattern)}"`);
  } else {
    attrs.push('w:val="clear"');
  }

  // Color (pattern color)
  if (color) {
    const patternColor: ExhaustiveColorValue = color;
    const {
      rgb,
      auto,
      themeColor: _themeColor,
      themeTint: _themeTint,
      themeShade: _themeShade,
    } = patternColor;
    if (rgb && isValidHexColor(rgb)) {
      attrs.push(`w:color="${escapeXmlAttribute(rgb)}"`);
    } else if (auto) {
      attrs.push('w:color="auto"');
    }
  }

  // Fill (background color)
  if (fill) {
    const fillColor: ExhaustiveColorValue = fill;
    const { rgb, auto, themeColor, themeTint, themeShade } = fillColor;
    if (rgb && isValidHexColor(rgb)) {
      attrs.push(`w:fill="${escapeXmlAttribute(rgb)}"`);
    } else if (auto) {
      attrs.push('w:fill="auto"');
    }
    if (themeColor) {
      attrs.push(`w:themeFill="${escapeXmlAttribute(themeColor)}"`);
    }
    if (themeTint) {
      attrs.push(`w:themeFillTint="${escapeXmlAttribute(themeTint)}"`);
    }
    if (themeShade) {
      attrs.push(`w:themeFillShade="${escapeXmlAttribute(themeShade)}"`);
    }
  }

  return attrs.length === 0 ? "" : `<w:shd ${attrs.join(" ")}/>`;
}

// ============================================================================
// TEXT FORMATTING SERIALIZATION
// ============================================================================

/** `CT_OnOff`: present means on, and an explicit off is not an absent one. */
const onOff = (name: string, value: boolean | undefined): string => {
  if (value === undefined) {
    return "";
  }
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
};

const numberTag = (name: string, value: number | undefined): string =>
  value === undefined ? "" : `<w:${name} w:val="${intAttr(value)}"/>`;

const serializeFontFamily = (fontFamily: FontFamily | undefined): string => {
  if (!fontFamily) {
    return "";
  }
  const exhaustiveFontFamily: ExhaustiveFontFamily = fontFamily;
  const {
    ascii,
    hAnsi,
    eastAsia,
    cs: complexScript,
    hint,
    asciiTheme,
    hAnsiTheme,
    eastAsiaTheme,
    csTheme,
  } = exhaustiveFontFamily;
  const attrs: string[] = [];
  if (ascii) {
    attrs.push(`w:ascii="${escapeXmlAttribute(ascii)}"`);
  }
  if (hAnsi) {
    attrs.push(`w:hAnsi="${escapeXmlAttribute(hAnsi)}"`);
  }
  if (eastAsia) {
    attrs.push(`w:eastAsia="${escapeXmlAttribute(eastAsia)}"`);
  }
  if (complexScript) {
    attrs.push(`w:cs="${escapeXmlAttribute(complexScript)}"`);
  }
  if (hint) {
    attrs.push(`w:hint="${escapeXmlAttribute(hint)}"`);
  }
  if (asciiTheme) {
    attrs.push(`w:asciiTheme="${escapeXmlAttribute(asciiTheme)}"`);
  }
  if (hAnsiTheme) {
    attrs.push(`w:hAnsiTheme="${escapeXmlAttribute(hAnsiTheme)}"`);
  }
  if (eastAsiaTheme) {
    attrs.push(`w:eastAsiaTheme="${escapeXmlAttribute(eastAsiaTheme)}"`);
  }
  if (csTheme) {
    // OOXML spells this attribute all-lowercase (`w:cstheme`), unlike its
    // camelCase siblings above; the parser reads `w:cstheme`, so emitting
    // `w:csTheme` would silently drop the CS theme font on round-trip.
    attrs.push(`w:cstheme="${escapeXmlAttribute(csTheme)}"`);
  }
  return attrs.length === 0 ? "" : `<w:rFonts ${attrs.join(" ")}/>`;
};

const serializeUnderline = (underline: Underline | undefined): string => {
  if (!underline) {
    return "";
  }
  const exhaustiveUnderline: ExhaustiveUnderline = underline;
  const { style, color } = exhaustiveUnderline;
  const attrs: string[] = [`w:val="${style}"`];
  if (color) {
    const exhaustiveColor: ExhaustiveColorValue = color;
    const { rgb, themeColor, themeTint, themeShade, auto: _auto } = exhaustiveColor;
    if (rgb && isValidHexColor(rgb)) {
      attrs.push(`w:color="${escapeXmlAttribute(rgb)}"`);
    }
    if (themeColor) {
      attrs.push(`w:themeColor="${escapeXmlAttribute(themeColor)}"`);
    }
    if (themeTint) {
      attrs.push(`w:themeTint="${escapeXmlAttribute(themeTint)}"`);
    }
    if (themeShade) {
      attrs.push(`w:themeShade="${escapeXmlAttribute(themeShade)}"`);
    }
  }
  return `<w:u ${attrs.join(" ")}/>`;
};

const serializeLanguage = (language: Language | undefined): string => {
  if (!language) {
    return "";
  }
  const exhaustiveLanguage: ExhaustiveLanguage = language;
  const { val, eastAsia, bidi } = exhaustiveLanguage;
  const attrs: string[] = [];
  if (val) {
    attrs.push(`w:val="${escapeXmlAttribute(val)}"`);
  }
  if (eastAsia) {
    attrs.push(`w:eastAsia="${escapeXmlAttribute(eastAsia)}"`);
  }
  if (bidi) {
    attrs.push(`w:bidi="${escapeXmlAttribute(bidi)}"`);
  }
  return attrs.length === 0 ? "" : `<w:lang ${attrs.join(" ")}/>`;
};

/**
 * The single `w:rPr` writer.
 *
 * A run, the paragraph mark, a style, a numbering level and the snapshot
 * inside either kind of `w:rPrChange` all write their run properties through
 * here, so none of them can grow its own order or its own idea of which
 * children exist. `EG_RPrBase` is a sequence and a consumer refuses a `w:rPr`
 * whose children are out of it, so the order is read from the generated
 * declared-child list rather than from the order of the statements below —
 * the restatement is what drifted, and it is why folio wrote `w:vanish`
 * before `w:noProof` while the schema declares the reverse.
 *
 * @param owned children a sibling record holds; see {@link OwnedRunPropertyChild}
 */
export function serializeTextFormatting(
  input: ExhaustiveTextFormatting | undefined,
  owned: readonly OwnedRunPropertyChild[] = [],
): string {
  const formatting: ExhaustiveTextFormatting = input ?? {};

  const {
    bold,
    boldCs,
    italic,
    italicCs,
    underline,
    strike,
    doubleStrike,
    vertAlign,
    smallCaps,
    allCaps,
    hidden,
    noProof,
    color,
    highlight,
    shading,
    fontSize,
    fontSizeCs,
    fontFamily,
    language,
    spacing,
    position,
    scale,
    kerning,
    effect,
    emphasisMark,
    emboss,
    imprint,
    outline,
    shadow,
    rtl,
    cs,
    styleId,
    preserved,
  } = formatting;

  // Emit valid OOXML named highlight colors through `w:highlight`, `none`
  // included, because that is what cancels an inherited highlight. A custom
  // color has no `w:highlight` spelling and falls back to `w:shd`, which sits
  // at its own place in the sequence.
  const namedHighlight = highlight && VALID_HIGHLIGHT_COLORS.has(highlight) ? highlight : undefined;
  const customHighlightHex =
    highlight && !namedHighlight && !shading ? highlight.replace(/^#/u, "") : "";
  const customHighlightShadingXml = /^[0-9a-fA-F]{6}$/u.test(customHighlightHex)
    ? `<w:shd w:val="clear" w:color="auto" w:fill="${customHighlightHex}"/>`
    : "";

  const parts = serializeSequenceChildren({
    container: "run-properties",
    modelled: [
      ...owned,
      ["rStyle", styleId ? `<w:rStyle w:val="${escapeXmlAttribute(styleId)}"/>` : ""],
      ["rFonts", serializeFontFamily(fontFamily)],
      ["b", onOff("b", bold)],
      ["bCs", onOff("bCs", boldCs)],
      ["i", onOff("i", italic)],
      ["iCs", onOff("iCs", italicCs)],
      ["caps", onOff("caps", allCaps)],
      ["smallCaps", onOff("smallCaps", smallCaps)],
      ["strike", onOff("strike", strike)],
      ["dstrike", onOff("dstrike", doubleStrike)],
      ["outline", onOff("outline", outline)],
      ["shadow", onOff("shadow", shadow)],
      ["emboss", onOff("emboss", emboss)],
      ["imprint", onOff("imprint", imprint)],
      ["noProof", onOff("noProof", noProof)],
      ["vanish", onOff("vanish", hidden)],
      ["color", serializeColorElement(color)],
      ["spacing", numberTag("spacing", spacing)],
      ["w", numberTag("w", roundHorizontalScalePercentForSerialization(scale))],
      ["kern", numberTag("kern", kerning)],
      ["position", numberTag("position", position)],
      ["sz", numberTag("sz", fontSize)],
      ["szCs", numberTag("szCs", fontSizeCs)],
      ["highlight", namedHighlight ? `<w:highlight w:val="${namedHighlight}"/>` : ""],
      ["u", serializeUnderline(underline)],
      ["effect", effect ? `<w:effect w:val="${effect}"/>` : ""],
      ["shd", serializeShading(shading) || customHighlightShadingXml],
      ["vertAlign", vertAlign ? `<w:vertAlign w:val="${vertAlign}"/>` : ""],
      ["rtl", onOff("rtl", rtl)],
      ["cs", onOff("cs", cs)],
      ["em", emphasisMark ? `<w:em w:val="${emphasisMark}"/>` : ""],
      ["lang", serializeLanguage(language)],
    ],
    preserved,
  });

  return parts.length === 0 ? "" : `<w:rPr>${parts.join("")}</w:rPr>`;
}
