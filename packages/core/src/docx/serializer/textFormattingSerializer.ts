import type {
  ColorValue,
  ExhaustiveFields,
  ShadingProperties,
  TextFormatting,
} from "../../types/document";
import { HIGHLIGHT_COLOR_VALUES } from "../../types/documentEnumValues";
import { isValidHexColor } from "../../utils/colorResolver";
import { roundHorizontalScalePercentForSerialization } from "../../utils/horizontalScale";
import { escapeXml, intAttr } from "./xmlUtils";

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
  | "styleId";
type ExhaustiveTextFormatting = ExhaustiveFields<TextFormatting, ClassifiedTextFormattingField>;

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
    attrs.push(`w:val="${escapeXml(rgb)}"`);
  }

  if (themeColor) {
    attrs.push(`w:themeColor="${escapeXml(themeColor)}"`);
  }

  if (themeTint) {
    attrs.push(`w:themeTint="${escapeXml(themeTint)}"`);
  }

  if (themeShade) {
    attrs.push(`w:themeShade="${escapeXml(themeShade)}"`);
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
    attrs.push(`w:val="${escapeXml(pattern)}"`);
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
      attrs.push(`w:color="${escapeXml(rgb)}"`);
    } else if (auto) {
      attrs.push('w:color="auto"');
    }
  }

  // Fill (background color)
  if (fill) {
    const fillColor: ExhaustiveColorValue = fill;
    const { rgb, auto, themeColor, themeTint, themeShade } = fillColor;
    if (rgb && isValidHexColor(rgb)) {
      attrs.push(`w:fill="${escapeXml(rgb)}"`);
    } else if (auto) {
      attrs.push('w:fill="auto"');
    }
    if (themeColor) {
      attrs.push(`w:themeFill="${escapeXml(themeColor)}"`);
    }
    if (themeTint) {
      attrs.push(`w:themeFillTint="${escapeXml(themeTint)}"`);
    }
    if (themeShade) {
      attrs.push(`w:themeFillShade="${escapeXml(themeShade)}"`);
    }
  }

  return attrs.length === 0 ? "" : `<w:shd ${attrs.join(" ")}/>`;
}

// ============================================================================
// TEXT FORMATTING SERIALIZATION
// ============================================================================

/**
 * Serialize text formatting properties to w:rPr XML
 */
export function serializeTextFormatting(input: ExhaustiveTextFormatting | undefined): string {
  if (!input) {
    return "";
  }

  const formatting: ExhaustiveTextFormatting = input;

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
  } = formatting;

  const parts: string[] = [];

  // Style reference (must be first)
  if (styleId) {
    parts.push(`<w:rStyle w:val="${escapeXml(styleId)}"/>`);
  }

  // Font family (w:rFonts)
  if (fontFamily) {
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
    const fontAttrs: string[] = [];
    if (ascii) {
      fontAttrs.push(`w:ascii="${escapeXml(ascii)}"`);
    }
    if (hAnsi) {
      fontAttrs.push(`w:hAnsi="${escapeXml(hAnsi)}"`);
    }
    if (eastAsia) {
      fontAttrs.push(`w:eastAsia="${escapeXml(eastAsia)}"`);
    }
    if (complexScript) {
      fontAttrs.push(`w:cs="${escapeXml(complexScript)}"`);
    }
    if (hint) {
      fontAttrs.push(`w:hint="${escapeXml(hint)}"`);
    }
    if (asciiTheme) {
      fontAttrs.push(`w:asciiTheme="${escapeXml(asciiTheme)}"`);
    }
    if (hAnsiTheme) {
      fontAttrs.push(`w:hAnsiTheme="${escapeXml(hAnsiTheme)}"`);
    }
    if (eastAsiaTheme) {
      fontAttrs.push(`w:eastAsiaTheme="${escapeXml(eastAsiaTheme)}"`);
    }
    if (csTheme) {
      // OOXML spells this attribute all-lowercase (`w:cstheme`), unlike its
      // camelCase siblings above; the parser reads `w:cstheme`, so emitting
      // `w:csTheme` would silently drop the CS theme font on round-trip.
      fontAttrs.push(`w:cstheme="${escapeXml(csTheme)}"`);
    }
    if (fontAttrs.length > 0) {
      parts.push(`<w:rFonts ${fontAttrs.join(" ")}/>`);
    }
  }

  // Bold
  if (bold === true) {
    parts.push("<w:b/>");
  } else if (bold === false) {
    parts.push('<w:b w:val="0"/>');
  }

  if (boldCs === true) {
    parts.push("<w:bCs/>");
  } else if (boldCs === false) {
    parts.push('<w:bCs w:val="0"/>');
  }

  // Italic
  if (italic === true) {
    parts.push("<w:i/>");
  } else if (italic === false) {
    parts.push('<w:i w:val="0"/>');
  }

  if (italicCs === true) {
    parts.push("<w:iCs/>");
  } else if (italicCs === false) {
    parts.push('<w:iCs w:val="0"/>');
  }

  // Caps
  if (allCaps === true) {
    parts.push("<w:caps/>");
  } else if (allCaps === false) {
    parts.push('<w:caps w:val="0"/>');
  }

  if (smallCaps === true) {
    parts.push("<w:smallCaps/>");
  } else if (smallCaps === false) {
    parts.push('<w:smallCaps w:val="0"/>');
  }

  // Strike
  if (strike === true) {
    parts.push("<w:strike/>");
  } else if (strike === false) {
    parts.push('<w:strike w:val="0"/>');
  }

  if (doubleStrike === true) {
    parts.push("<w:dstrike/>");
  } else if (doubleStrike === false) {
    parts.push('<w:dstrike w:val="0"/>');
  }

  // Outline
  if (outline === true) {
    parts.push("<w:outline/>");
  } else if (outline === false) {
    parts.push('<w:outline w:val="0"/>');
  }

  // Shadow
  if (shadow === true) {
    parts.push("<w:shadow/>");
  } else if (shadow === false) {
    parts.push('<w:shadow w:val="0"/>');
  }

  // Emboss
  if (emboss === true) {
    parts.push("<w:emboss/>");
  } else if (emboss === false) {
    parts.push('<w:emboss w:val="0"/>');
  }

  // Imprint
  if (imprint === true) {
    parts.push("<w:imprint/>");
  } else if (imprint === false) {
    parts.push('<w:imprint w:val="0"/>');
  }

  // Hidden
  if (hidden === true) {
    parts.push("<w:vanish/>");
  } else if (hidden === false) {
    parts.push('<w:vanish w:val="0"/>');
  }
  if (noProof === true) {
    parts.push("<w:noProof/>");
  } else if (noProof === false) {
    parts.push('<w:noProof w:val="0"/>');
  }

  // Color
  const colorXml = serializeColorElement(color);
  if (colorXml) {
    parts.push(colorXml);
  }

  // Spacing
  if (spacing !== undefined) {
    parts.push(`<w:spacing w:val="${intAttr(spacing)}"/>`);
  }

  // Scale (w:w)
  const horizontalScale = roundHorizontalScalePercentForSerialization(scale);
  if (horizontalScale !== undefined) {
    parts.push(`<w:w w:val="${intAttr(horizontalScale)}"/>`);
  }

  // Kerning
  if (kerning !== undefined) {
    parts.push(`<w:kern w:val="${intAttr(kerning)}"/>`);
  }

  // Position
  if (position !== undefined) {
    parts.push(`<w:position w:val="${intAttr(position)}"/>`);
  }

  // Font size
  if (fontSize !== undefined) {
    parts.push(`<w:sz w:val="${intAttr(fontSize)}"/>`);
  }

  if (fontSizeCs !== undefined) {
    parts.push(`<w:szCs w:val="${intAttr(fontSizeCs)}"/>`);
  }

  // Highlight — emit valid OOXML named colors via w:highlight, including
  // `none`, which explicitly cancels an inherited highlight. A custom color
  // falls back to w:shd at that property's later CT_RPr position.
  let customHighlightShadingXml = "";
  if (highlight) {
    if (VALID_HIGHLIGHT_COLORS.has(highlight)) {
      parts.push(`<w:highlight w:val="${highlight}"/>`);
    } else if (!shading) {
      // Custom color not in OOXML predefined set — use w:shd as fallback.
      // Only emit if value looks like a valid hex color.
      const hex = highlight.replace(/^#/u, "");
      if (/^[0-9a-fA-F]{6}$/u.test(hex)) {
        customHighlightShadingXml = `<w:shd w:val="clear" w:color="auto" w:fill="${hex}"/>`;
      }
    }
  }

  // Underline
  if (underline) {
    const exhaustiveUnderline: ExhaustiveUnderline = underline;
    const { style, color: underlineColor } = exhaustiveUnderline;
    const uAttrs: string[] = [`w:val="${style}"`];
    if (underlineColor) {
      const exhaustiveUnderlineColor: ExhaustiveColorValue = underlineColor;
      const { rgb, themeColor, themeTint, themeShade, auto: _auto } = exhaustiveUnderlineColor;
      if (rgb && isValidHexColor(rgb)) {
        uAttrs.push(`w:color="${escapeXml(rgb)}"`);
      }
      if (themeColor) {
        uAttrs.push(`w:themeColor="${escapeXml(themeColor)}"`);
      }
      if (themeTint) {
        uAttrs.push(`w:themeTint="${escapeXml(themeTint)}"`);
      }
      if (themeShade) {
        uAttrs.push(`w:themeShade="${escapeXml(themeShade)}"`);
      }
    }
    parts.push(`<w:u ${uAttrs.join(" ")}/>`);
  }

  // Effect
  if (effect) {
    parts.push(`<w:effect w:val="${effect}"/>`);
  }

  // Shading
  const shadingXml = serializeShading(shading) || customHighlightShadingXml;
  if (shadingXml) {
    parts.push(shadingXml);
  }

  // Vertical alignment
  if (vertAlign) {
    parts.push(`<w:vertAlign w:val="${vertAlign}"/>`);
  }

  // RTL and CS
  if (rtl === true) {
    parts.push("<w:rtl/>");
  } else if (rtl === false) {
    // Preserve explicit overrides like <w:rtl w:val="0"/> that disable
    // inherited paragraph/style RTL.
    parts.push('<w:rtl w:val="0"/>');
  }

  if (cs === true) {
    parts.push("<w:cs/>");
  } else if (cs === false) {
    parts.push('<w:cs w:val="0"/>');
  }

  // Emphasis mark
  if (emphasisMark) {
    parts.push(`<w:em w:val="${emphasisMark}"/>`);
  }

  if (language) {
    const exhaustiveLanguage: ExhaustiveLanguage = language;
    const { val, eastAsia, bidi } = exhaustiveLanguage;
    const languageAttrs: string[] = [];
    if (val) {
      languageAttrs.push(`w:val="${escapeXml(val)}"`);
    }
    if (eastAsia) {
      languageAttrs.push(`w:eastAsia="${escapeXml(eastAsia)}"`);
    }
    if (bidi) {
      languageAttrs.push(`w:bidi="${escapeXml(bidi)}"`);
    }
    if (languageAttrs.length > 0) {
      parts.push(`<w:lang ${languageAttrs.join(" ")}/>`);
    }
  }

  return parts.length === 0 ? "" : `<w:rPr>${parts.join("")}</w:rPr>`;
}
