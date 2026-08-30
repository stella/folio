/**
 * ProseMirror to FlowBlock Converter
 *
 * Converts a ProseMirror document into FlowBlock[] for the layout engine.
 * Tracks pmStart/pmEnd positions for click-to-position mapping.
 */

import type { Node as PMNode, Mark } from "prosemirror-model";

import { convertBulletToUnicode } from "../../docx/bulletMarkers";
import { resolveDocumentGridLinePitch } from "../../docx/documentGrid";
import { getFontAlternate, type FontAlternates } from "../../fonts/fontAlternates";
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TableRow,
  TableCell,
  CellBorders,
  BorderStyle,
  ImageBlock,
  TextBoxBlock,
  PageBreakBlock,
  ColumnBreakBlock,
  SectionBreakBlock,
  Run,
  TextRun,
  TabRun,
  ImageRun,
  FieldRun,
  RunFormatting,
  ParagraphAttrs,
  SdtGroup,
  TabStop,
  FloatingTablePosition,
} from "../../layout-engine/types";
import { setHyperlinkInstanceIndex } from "../../layout-engine/measure/hyperlinkInstance";
import { setTextBoxGroupId } from "../../layout-engine/textBoxGroup";
import { setParagraphFrame } from "../../layout-engine/paragraphFrame";
import { DEFAULT_TEXTBOX_MARGINS, DEFAULT_TEXTBOX_WIDTH } from "../../layout-engine/types";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import { getColumns } from "../sectionColumns";
import {
  expectBlockSdtAttrs,
  expectCharacterStyleMarkAttrs,
  expectCharacterSpacingMarkAttrs,
  expectCommentMarkAttrs,
  expectEmphasisMarkAttrs,
  expectFieldAttrs,
  expectFontFamilyMarkAttrs,
  expectLanguageMarkAttrs,
  expectFontSizeMarkAttrs,
  expectFootnoteRefMarkAttrs,
  expectHardBreakAttrs,
  expectHighlightMarkAttrs,
  expectHyperlinkMarkAttrs,
  expectImageAttrs,
  expectMathAttrs,
  expectParagraphAttrs,
  expectRunFormattingOverrideMarkAttrs,
  expectRunShadingMarkAttrs,
  expectSymbolAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
  expectTextBoxAttrs,
  expectTextColorMarkAttrs,
  expectTextEffectMarkAttrs,
  expectTrackedChangeMarkAttrs,
  expectUnderlineMarkAttrs,
} from "../../prosemirror/attrs";
import { autospacingMatchesBase } from "../../prosemirror/autospacingBase";
import { runShadingAttrsToShading } from "../../prosemirror/conversion/runShadingMark";
import { directionToBidi } from "../../prosemirror/paragraphDirection";
import { cascadeStyleTextFormatting } from "../../prosemirror/styles/styleToggleCascade";
import { getPageNumbering } from "../../paged-layout/sectionGeometry";
import type { RunFormattingOverrideAttrs } from "../../prosemirror/schema/marks";
import type {
  ImageAttrs,
  ParagraphPropertyChangeAttrs,
  ParagraphAttrs as PMParagraphAttrs,
} from "../../prosemirror/schema/nodes";
import { assertValidProseMirrorDocument } from "../../prosemirror/validation";
import {
  advanceListMarker,
  advanceVisibleListMarker,
  cloneListCounterState,
  type ListCounterState,
  type ListCounterStreams,
} from "../../prosemirror/listMarker";
import { resolveNumberedRefFields } from "../../prosemirror/numberedRefFields";
import type { ColorValue, Theme, SectionProperties, TextFormatting } from "../../types/document";
import { normalizeShapeTextAnchor } from "../../types/documentEnumValues";
import { resolveColor, resolveHighlightToCss } from "../../utils/colorResolver";
import { resolveThemeFont } from "../../utils/fontResolver";
import { resolveShadingFill } from "../../utils/formatToStyle";
import { decodeOoxmlSymbolCharacter } from "../../utils/ooxmlSymbol";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import {
  AUTO_PARAGRAPH_SPACING_PX,
  pointsToPixels,
  halfPointsToPixels,
  halfPointsToPoints,
} from "../../utils/units";
import { groupParagraphFrames } from "./paragraphFrames";

export { formatCounter, resolveListTemplate } from "../../prosemirror/listMarker";

/**
 * Options for the conversion.
 */
export type ToFlowBlocksOptions = {
  /** Default font family. */
  defaultFont?: string;
  /** Default font size in points. */
  defaultSize?: number;
  /** Theme for resolving theme colors. */
  theme?: Theme | null;
  /** Document-scoped OOXML primary-font to `w:altName` lookup. */
  fontAlternates?: FontAlternates;
  /** Page content height in pixels (pageHeight - marginTop - marginBottom). Images taller than this are scaled down to fit. */
  pageContentHeight?: number;
  /** Shared list counters for nested containers. */
  listCounters?: Map<number, number[]>;
  /** Latest concrete counters by abstract numbering definition. */
  listAbstractCounters?: Map<number, number[]>;
  /** Shared startOverride state for nested containers. */
  listSeenNumIds?: Set<string>;
  /**
   * Parallel counter state for the "original" (pre-revision) document, used to
   * number tracked-deletion list items. Word numbers inserted and deleted list
   * runs as if they never coexist: insertions get final-document numbering,
   * deletions keep their original numbering. Without a separate stream a deleted
   * item continues the counter of the inserted item before it (a, b → c, d, e
   * instead of a, b and a, b, c). Normal items advance both streams.
   */
  originalListCounters?: Map<number, number[]>;
  /** Latest concrete original-stream counters by abstract numbering definition. */
  originalListAbstractCounters?: Map<number, number[]>;
  /** Original-stream startOverride state. */
  originalListSeenNumIds?: Set<string>;
  /**
   * Document-wide `w:defaultTabStop` (§17.6.13) in twips. Stamped onto
   * every paragraph block so paragraph-local layout helpers (list marker
   * tab-stop math) can read it without taking a `Document` reference.
   * Defaults to the OOXML 720-twip value when absent.
   */
  defaultTabStopTwips?: number;
  /** Document-wide custom Word line-breaking settings. */
  lineBreakRules?: {
    noLineBreaksBefore?: { language?: string; characters: string };
    noLineBreaksAfter?: { language?: string; characters: string };
    useLegacyEthiopicAmharicRules?: boolean;
  };
  /** Document-generation policy for justified line fitting. */
  justificationCompatibility?: NonNullable<ParagraphAttrs["justificationCompatibility"]>;
  /** Document-wide automatic hyphenation policy. */
  automaticHyphenation?: NonNullable<ParagraphAttrs["automaticHyphenation"]>;
  /** Line pitch for the final body section, whose properties live outside the PM body. */
  finalSectionDocumentGridLinePitchTwips?: number;
};

type FlowConversionOptions = ToFlowBlocksOptions & {
  listCounterStreams: ListCounterStreams;
  numberedRefResults?: ReadonlyMap<PMNode, string>;
};

const DEFAULT_FONT = "Calibri";
const DEFAULT_TABLE_CELL_MARGIN_TWIPS = {
  top: 0,
  right: 108,
  bottom: 0,
  left: 108,
} as const;
type TablePaddingSide = keyof typeof DEFAULT_TABLE_CELL_MARGIN_TWIPS;
const DEFAULT_BLACK_TEXT_COLOR_VALUES = new Set(["000000", "000"]);

function normalizeResolvedTextColor(color: string): string {
  return color.trim().toLowerCase().replace(/^#/u, "");
}

function isDefaultBlackResolvedTextColor(color: string): boolean {
  return DEFAULT_BLACK_TEXT_COLOR_VALUES.has(normalizeResolvedTextColor(color));
}

function areResolvedTextColorsEqual(left: string, right: string): boolean {
  return normalizeResolvedTextColor(left) === normalizeResolvedTextColor(right);
}

/**
 * Constrain image dimensions to fit within the page content area.
 * Scales proportionally if height exceeds pageContentHeight.
 */
function constrainImageToPage(
  width: number,
  height: number,
  pageContentHeight: number | undefined,
): { width: number; height: number } {
  if (!pageContentHeight || height <= pageContentHeight) {
    return { width, height };
  }
  const scale = pageContentHeight / height;
  return { width: Math.round(width * scale), height: pageContentHeight };
}

const DEFAULT_SIZE = 11; // points (Word 2007+ default)

/**
 * Convert twips to pixels (1 twip = 1/1440 inch, 1 inch = 96 CSS px).
 * No rounding — precision prevents cumulative layout drift across paragraphs.
 */
function twipsToPixels(twips: number): number {
  return (twips / 1440) * 96;
}

/**
 * Generate a unique block ID.
 */
let blockIdCounter = 0;
function nextBlockId(): string {
  return `block-${++blockIdCounter}`;
}

function applyMarkerAllCaps(marker: string | null, allCaps: boolean | undefined): string | null {
  if (marker === null || !allCaps) {
    return marker;
  }
  return marker.toLocaleUpperCase();
}

function computeListMarker(pmAttrs: PMParagraphAttrs, state: ListCounterState): string | null {
  return advanceListMarker(pmAttrs, state);
}

/**
 * Reset the block ID counter (useful for testing).
 */
export function resetBlockIdCounter(): void {
  blockIdCounter = 0;
}

/**
 * Extract run formatting from ProseMirror marks.
 */
function extractRunFormatting(
  marks: readonly Mark[],
  theme?: Theme | null,
  fontAlternates?: FontAlternates,
): RunFormatting {
  const formatting: RunFormatting = {};
  let hasNoteRef = false;

  for (const mark of marks) {
    switch (mark.type.name) {
      case "bold":
        formatting.bold = true;
        break;

      case "italic":
        formatting.italic = true;
        break;

      case "underline": {
        const attrs = expectUnderlineMarkAttrs(mark);
        if (attrs.style || attrs.color) {
          const underlineObj: { style?: string; color?: string } = {};
          if (attrs.style) {
            underlineObj.style = attrs.style;
          }
          if (attrs.color) {
            underlineObj.color = resolveColor(attrs.color, theme);
          }
          formatting.underline = underlineObj;
        } else {
          formatting.underline = true;
        }
        break;
      }

      case "strike":
        formatting.strike = true;
        break;

      case "textColor": {
        const attrs = expectTextColorMarkAttrs(mark);
        if (attrs.themeColor || attrs.rgb) {
          const colorArg: ColorValue = {};
          if (attrs.rgb) {
            colorArg.rgb = attrs.rgb;
          }
          if (attrs.themeColor) {
            colorArg.themeColor = attrs.themeColor;
          }
          if (attrs.themeTint) {
            colorArg.themeTint = attrs.themeTint;
          }
          if (attrs.themeShade) {
            colorArg.themeShade = attrs.themeShade;
          }
          if (!isAutomaticTextColorValue(colorArg)) {
            formatting.color = resolveColor(colorArg, theme);
            formatting.textColorSource = "direct";
          }
        }
        break;
      }

      case "highlight":
        formatting.highlight = resolveHighlightToCss(expectHighlightMarkAttrs(mark).color);
        break;

      case "runShading": {
        const shadingCss = resolveShadingFill(
          runShadingAttrsToShading(expectRunShadingMarkAttrs(mark)),
          theme,
        );
        if (shadingCss) {
          formatting.shading = shadingCss;
        }
        break;
      }

      case "fontSize": {
        const attrs = expectFontSizeMarkAttrs(mark);
        // Convert half-points to points
        formatting.fontSize = attrs.size / 2;
        break;
      }

      case "fontFamily": {
        const attrs = expectFontFamilyMarkAttrs(mark);
        const font = resolveWesternThemeFont(attrs, theme);
        if (font) {
          formatting.fontFamily = font;
          const alternate = getFontAlternate(font, fontAlternates);
          if (alternate) {
            formatting.alternateFontFamily = alternate;
          }
        }
        const eastAsiaFont = resolveEastAsiaThemeFont(attrs, theme);
        if (eastAsiaFont) {
          formatting.eastAsiaFontFamily = eastAsiaFont;
          const alternate = getFontAlternate(eastAsiaFont, fontAlternates);
          if (alternate) {
            formatting.eastAsiaAlternateFontFamily = alternate;
          }
        }
        const complexScriptFont = resolveComplexScriptThemeFont(attrs, theme);
        if (complexScriptFont) {
          formatting.complexScriptFontFamily = complexScriptFont;
          const alternate = getFontAlternate(complexScriptFont, fontAlternates);
          if (alternate) {
            formatting.complexScriptAlternateFontFamily = alternate;
          }
        }
        break;
      }

      case "language": {
        const attrs = expectLanguageMarkAttrs(mark);
        formatting.language = {
          ...(attrs.val ? { val: attrs.val } : {}),
          ...(attrs.eastAsia ? { eastAsia: attrs.eastAsia } : {}),
          ...(attrs.bidi ? { bidi: attrs.bidi } : {}),
        };
        break;
      }

      case "characterSpacing": {
        const attrs = expectCharacterSpacingMarkAttrs(mark);
        if (attrs.spacing !== undefined) {
          formatting.letterSpacing = twipsToPixels(attrs.spacing);
        }
        if (attrs.position !== undefined && attrs.position !== 0) {
          formatting.positionPx = halfPointsToPixels(attrs.position);
        }
        const horizontalScale = normalizeHorizontalScalePercent(attrs.scale);
        if (horizontalScale !== undefined) {
          formatting.horizontalScale = horizontalScale;
        }
        if (attrs.kerning !== undefined && attrs.kerning > 0) {
          formatting.kerningMinPt = halfPointsToPoints(attrs.kerning);
        }
        break;
      }

      case "allCaps":
        formatting.allCaps = true;
        break;

      case "smallCaps":
        formatting.smallCaps = true;
        break;

      case "emboss":
        formatting.emboss = true;
        break;

      case "imprint":
        formatting.imprint = true;
        break;

      case "hidden":
        // eigenpal #424 (w:vanish gap 9): mark surfaces RunFormatting.hidden
        // so the painter can apply the dimmed dotted-underline treatment.
        formatting.hidden = true;
        break;

      case "textShadow":
        formatting.textShadow = true;
        break;

      case "textOutline":
        formatting.textOutline = true;
        break;

      case "rtl":
        formatting.rtl = true;
        break;

      case "textEffect":
        // The textEffect mark schema rejects "none"; only animated variants
        // ever reach this branch.
        formatting.textEffect = expectTextEffectMarkAttrs(mark).effect;
        break;

      case "runFormattingOverride":
        applyRunFormattingOverrides(formatting, expectRunFormattingOverrideMarkAttrs(mark));
        break;

      case "emphasisMark": {
        formatting.emphasisMark = expectEmphasisMarkAttrs(mark).type ?? "dot";
        break;
      }

      case "superscript":
        formatting.superscript = true;
        break;

      case "subscript":
        formatting.subscript = true;
        break;

      case "hyperlink": {
        const attrs = expectHyperlinkMarkAttrs(mark);
        const link: RunFormatting["hyperlink"] & object = {
          href: attrs.href,
        };
        if (attrs.tooltip !== undefined) {
          link.tooltip = attrs.tooltip;
        }
        if (attrs._docxHyperlinkIndex !== undefined) {
          setHyperlinkInstanceIndex(link, attrs._docxHyperlinkIndex);
        }
        formatting.hyperlink = link;
        break;
      }

      case "footnoteRef": {
        hasNoteRef = true;
        const attrs = expectFootnoteRefMarkAttrs(mark);
        if (attrs.vertAlign === "superscript") {
          formatting.superscript = true;
        }
        const id = typeof attrs.id === "string" ? Number.parseInt(attrs.id, 10) : attrs.id;
        if (attrs.noteType === "endnote") {
          formatting.endnoteRefId = id;
        } else {
          formatting.footnoteRefId = id;
        }
        break;
      }

      case "comment": {
        const commentId = expectCommentMarkAttrs(mark).commentId;
        if (commentId) {
          if (!formatting.commentIds) {
            formatting.commentIds = [];
          }
          formatting.commentIds.push(commentId);
        }
        break;
      }

      case "insertion": {
        const attrs = expectTrackedChangeMarkAttrs(mark);
        formatting.isInsertion = true;
        formatting.changeAuthor = attrs.author;
        if (attrs.date !== undefined) {
          formatting.changeDate = attrs.date;
        }
        formatting.changeRevisionId = attrs.revisionId;
        if (attrs.provenance === "suggested") {
          formatting.isSuggestion = true;
          if (attrs.suggestionId) {
            formatting.suggestionId = attrs.suggestionId;
          }
        }
        break;
      }

      case "deletion": {
        const attrs = expectTrackedChangeMarkAttrs(mark);
        formatting.isDeletion = true;
        formatting.changeAuthor = attrs.author;
        if (attrs.date !== undefined) {
          formatting.changeDate = attrs.date;
        }
        formatting.changeRevisionId = attrs.revisionId;
        if (attrs.provenance === "suggested") {
          formatting.isSuggestion = true;
          if (attrs.suggestionId) {
            formatting.suggestionId = attrs.suggestionId;
          }
        }
        break;
      }
      default:
        break;
    }
  }

  if (hasNoteRef && formatting.subscript) {
    delete formatting.superscript;
  }

  return formatting;
}

type ThemeFontAttributes = Omit<NonNullable<TextFormatting["fontFamily"]>, "asciiTheme"> & {
  /** The ProseMirror attr reader validates this against the canonical theme values. */
  asciiTheme?: string;
};

const resolveWesternThemeFont = (
  fontFamily: ThemeFontAttributes,
  theme?: Theme | null,
): string | undefined => {
  const themeRef = fontFamily.asciiTheme ?? fontFamily.hAnsiTheme;
  const themedFont = themeRef ? resolveThemeFont(themeRef, theme?.fontScheme) : null;
  return themedFont ?? fontFamily.ascii ?? fontFamily.hAnsi ?? undefined;
};

const resolveComplexScriptThemeFont = (
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

const resolveEastAsiaThemeFont = (
  fontFamily: ThemeFontAttributes,
  theme?: Theme | null,
): string | undefined => {
  const themedFont = fontFamily.eastAsiaTheme
    ? resolveThemeFont(fontFamily.eastAsiaTheme, theme?.fontScheme)
    : null;
  return themedFont ?? fontFamily.eastAsia ?? undefined;
};

function isAutomaticTextColorValue(color: ColorValue): boolean {
  const rgb = color.rgb?.trim().toLowerCase();
  return color.auto === true || rgb === "auto" || (!rgb && !color.themeColor);
}

function markDefaultBlackTextColorSource(
  formatting: RunFormatting,
  paraDefaults: RunFormatting,
): RunFormatting {
  if (
    formatting.textColorSource === "direct" ||
    formatting.color === undefined ||
    paraDefaults.color === undefined ||
    !isDefaultBlackResolvedTextColor(formatting.color) ||
    !areResolvedTextColorsEqual(formatting.color, paraDefaults.color)
  ) {
    return formatting;
  }

  return {
    ...formatting,
    textColorSource: "paragraphDefault",
  };
}

function mergeRunFormatting(paraDefaults: RunFormatting, formatting: RunFormatting): RunFormatting {
  const merged = {
    ...paraDefaults,
    ...markDefaultBlackTextColorSource(formatting, paraDefaults),
  };
  if (formatting.fontFamily !== undefined && formatting.alternateFontFamily === undefined) {
    delete merged.alternateFontFamily;
  }
  if (
    formatting.eastAsiaFontFamily !== undefined &&
    formatting.eastAsiaAlternateFontFamily === undefined
  ) {
    delete merged.eastAsiaAlternateFontFamily;
  }
  if (
    formatting.complexScriptFontFamily !== undefined &&
    formatting.complexScriptAlternateFontFamily === undefined
  ) {
    delete merged.complexScriptAlternateFontFamily;
  }
  if (merged.letterSpacing === 0) {
    delete merged.letterSpacing;
  }
  if (formatting.horizontalScale === 100 && paraDefaults.horizontalScale === undefined) {
    delete merged.horizontalScale;
  }
  return merged;
}

type ApplyCharacterStyleToggleFormattingOptions = {
  formatting: RunFormatting;
  marks: readonly Mark[];
  paraDefaults: RunFormatting;
};

/** Restore character-style toggle values that plain visual marks cannot represent. */
function applyCharacterStyleToggleFormatting({
  formatting,
  marks,
  paraDefaults,
}: ApplyCharacterStyleToggleFormattingOptions): void {
  const characterStyleMark = marks.find((mark) => mark.type.name === "characterStyle");
  if (!characterStyleMark) {
    return;
  }
  const styleRPr = expectCharacterStyleMarkAttrs(characterStyleMark)._styleRPr;
  if (!styleRPr) {
    return;
  }

  const effectiveStyleFormatting = cascadeStyleTextFormatting([
    {
      formatting: {
        bold: paraDefaults.bold ?? false,
        boldCs: paraDefaults.complexScriptBold ?? false,
        italic: paraDefaults.italic ?? false,
        italicCs: paraDefaults.complexScriptItalic ?? false,
      },
      type: "direct",
    },
    { formatting: styleRPr, type: "style" },
  ]).formatting;

  if (styleRPr.bold !== undefined && formatting.bold === undefined) {
    formatting.bold = effectiveStyleFormatting?.bold ?? false;
  }
  if (styleRPr.boldCs !== undefined && formatting.complexScriptBold === undefined) {
    formatting.complexScriptBold = effectiveStyleFormatting?.boldCs ?? false;
  }
  if (styleRPr.italic !== undefined && formatting.italic === undefined) {
    formatting.italic = effectiveStyleFormatting?.italic ?? false;
  }
  if (styleRPr.italicCs !== undefined && formatting.complexScriptItalic === undefined) {
    formatting.complexScriptItalic = effectiveStyleFormatting?.italicCs ?? false;
  }
  if (styleRPr.allCaps === true && formatting.allCaps === undefined) {
    formatting.allCaps = false;
  }
  if (styleRPr.emboss === true && formatting.emboss === undefined) {
    formatting.emboss = false;
  }
  if (styleRPr.imprint === true && formatting.imprint === undefined) {
    formatting.imprint = false;
  }
  if (styleRPr.outline === true && formatting.textOutline === undefined) {
    formatting.textOutline = false;
  }
  if (styleRPr.shadow === true && formatting.textShadow === undefined) {
    formatting.textShadow = false;
  }
  if (styleRPr.smallCaps === true && formatting.smallCaps === undefined) {
    formatting.smallCaps = false;
  }
  if (styleRPr.strike === true && formatting.strike === undefined) {
    formatting.strike = false;
  }
  if (styleRPr.hidden === true && formatting.hidden === undefined) {
    formatting.hidden = false;
  }
}

function applyRunFormattingOverrides(
  formatting: RunFormatting,
  attrs: RunFormattingOverrideAttrs,
): void {
  if (attrs.bold !== undefined) {
    formatting.bold = attrs.bold;
  }
  if (attrs.italic !== undefined) {
    formatting.italic = attrs.italic;
  }
  if (attrs.underline === "none") {
    formatting.underline = false;
  }
  if (attrs.strike !== undefined) {
    formatting.strike = attrs.strike;
  }
  if (attrs.allCaps !== undefined) {
    formatting.allCaps = attrs.allCaps;
  }
  if (attrs.smallCaps !== undefined) {
    formatting.smallCaps = attrs.smallCaps;
  }
  if (attrs.hidden !== undefined) {
    formatting.hidden = attrs.hidden;
  }
  if (attrs.emboss !== undefined) {
    formatting.emboss = attrs.emboss;
  }
  if (attrs.imprint !== undefined) {
    formatting.imprint = attrs.imprint;
  }
  if (attrs.shadow !== undefined) {
    formatting.textShadow = attrs.shadow;
  }
  if (attrs.outline !== undefined) {
    formatting.textOutline = attrs.outline;
  }
  if (attrs.rtl === false) {
    formatting.rtl = false;
  }
  if (attrs.boldCs !== undefined) {
    formatting.complexScriptBold = attrs.boldCs;
  }
  if (attrs.italicCs !== undefined) {
    formatting.complexScriptItalic = attrs.italicCs;
  }
  if (attrs.fontSizeCs !== undefined) {
    formatting.complexScriptFontSize = attrs.fontSizeCs / 2;
  }
  if (attrs.cs !== undefined) {
    formatting.forceComplexScript = attrs.cs;
  }
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

function paragraphRunDefaults(
  pmAttrs: PMParagraphAttrs,
  theme?: Theme | null,
  fontAlternates?: FontAlternates,
): RunFormatting {
  return textFormattingToRunFormatting(pmAttrs.defaultTextFormatting, theme, fontAlternates);
}

/**
 * Build an ImageRun from ProseMirror node attrs, applying conditional property assignment
 * to satisfy exactOptionalPropertyTypes.
 */
function buildImageRun(
  attrs: ImageAttrs,
  constrained: { width: number; height: number },
  pmStart: number,
  pmEnd: number,
  // Tracked-change attrs lifted off the image node's PM marks. eigenpal #641.
  trackedChange?: Pick<
    RunFormatting,
    "isInsertion" | "isDeletion" | "changeAuthor" | "changeDate" | "changeRevisionId"
  >,
): ImageRun {
  const run: ImageRun = {
    kind: "image",
    src: attrs.src,
    width: constrained.width,
    height: constrained.height,
    pmStart,
    pmEnd,
  };
  if (attrs.alt !== undefined) {
    run.alt = attrs.alt;
  }
  if (attrs.transform !== undefined) {
    run.transform = attrs.transform;
  }
  // eigenpal #424 (opacity render pipeline): copy opacity verbatim. PM
  // schema defaults `opacity` to `null`, which survives the typed cast on
  // ImageAttrs (`number | undefined`). Gate with `!= null` so the model
  // never carries the schema sentinel.
  if (attrs.opacity != null) {
    run.opacity = attrs.opacity;
  }
  if (attrs.wrapType !== undefined) {
    run.wrapType = attrs.wrapType;
  }
  if (attrs.displayMode !== undefined) {
    run.displayMode = attrs.displayMode;
  }
  if (attrs.cssFloat !== undefined) {
    run.cssFloat = attrs.cssFloat;
  }
  if (attrs.distTop !== undefined) {
    run.distTop = attrs.distTop;
  }
  if (attrs.distBottom !== undefined) {
    run.distBottom = attrs.distBottom;
  }
  if (attrs.distLeft !== undefined) {
    run.distLeft = attrs.distLeft;
  }
  if (attrs.distRight !== undefined) {
    run.distRight = attrs.distRight;
  }
  if (attrs._docxObjectPreview === true) {
    run.exactLineHeight = true;
  }
  // eigenpal #424: pass crop fractions through to the painter so it can
  // emit CSS clip-path. PM defaults are `null`; treat null as "not set".
  if (attrs.cropTop != null) {
    run.cropTop = attrs.cropTop;
  }
  if (attrs.cropRight != null) {
    run.cropRight = attrs.cropRight;
  }
  if (attrs.cropBottom != null) {
    run.cropBottom = attrs.cropBottom;
  }
  if (attrs.cropLeft != null) {
    run.cropLeft = attrs.cropLeft;
  }
  // eigenpal #1096: image borders are authored on the PM image attrs and
  // painted by layout-painter. PM defaults are null; treat null as absent.
  if (attrs.borderWidth != null) {
    run.borderWidth = attrs.borderWidth;
  }
  if (attrs.borderColor) {
    run.borderColor = attrs.borderColor;
  }
  if (attrs.borderStyle) {
    run.borderStyle = attrs.borderStyle;
  }
  if (attrs.position !== undefined) {
    run.position = attrs.position;
  }
  if (attrs.layoutInCell !== undefined) {
    run.layoutInCell = attrs.layoutInCell;
  }
  if (trackedChange?.isInsertion) {
    run.isInsertion = true;
  }
  if (trackedChange?.isDeletion) {
    run.isDeletion = true;
  }
  if (trackedChange?.changeAuthor !== undefined) {
    run.changeAuthor = trackedChange.changeAuthor;
  }
  if (trackedChange?.changeDate !== undefined) {
    run.changeDate = trackedChange.changeDate;
  }
  if (trackedChange?.changeRevisionId !== undefined) {
    run.changeRevisionId = trackedChange.changeRevisionId;
  }
  return run;
}

/**
 * In TOC paragraphs, strip the resolved Hyperlink character-style colour and
 * underline so the painter's link fallback doesn't fire. The PM doc keeps the
 * original marks so copy/paste out of a TOC still carries the Hyperlink
 * styling like Word does. Applies to both text and field runs — a TOC entry's
 * page number is a PAGEREF field inside the entry's hyperlink.
 *
 * Mutates `formatting` in place; cheaper than re-cloning per run.
 */
function stripTocHyperlinkStyle(formatting: RunFormatting): void {
  if (!formatting.hyperlink) {
    return;
  }
  formatting.hyperlink.noDefaultStyle = true;
  delete formatting.color;
  delete formatting.underline;
}

/**
 * Convert a paragraph node to runs.
 */
function paragraphToRuns(node: PMNode, startPos: number, _options: FlowConversionOptions): Run[] {
  const runs: Run[] = [];
  const offset = startPos + 1; // +1 for opening tag
  const theme = _options.theme;
  const fontAlternates = _options.fontAlternates;
  const pmAttrs = expectParagraphAttrs(node);
  const paraDefaults = paragraphRunDefaults(pmAttrs, theme, fontAlternates);
  const paragraphStyleId = pmAttrs.styleId;
  const inTocParagraph =
    pmAttrs._tableOfContentsLevel !== undefined ||
    tableOfContentsStyleLevel({ styleId: paragraphStyleId }) !== undefined;
  let leadingRenderedPageBreakPending = pmAttrs.renderedPageBreakBefore === true;

  // Single dispatcher for one inline PM child. Recurses on `sdt` so nested
  // content controls keep contributing runs at the right pmStart/pmEnd.
  // Used for both the top-level paragraph iteration and the descent into
  // SDT children — the previous SDT branch only handled text/hardBreak/
  // tab/image and silently dropped fields, math, and nested SDTs even
  // when the parser preserved them (see eigenpal #482).
  function pushRunsForChild(child: PMNode, childPos: number): void {
    if (child.type.name === "bookmarkBoundary") {
      return;
    }
    if (child.type.name === "renderedPageBreak") {
      if (leadingRenderedPageBreakPending) {
        leadingRenderedPageBreakPending = false;
        return;
      }
      runs.push({
        kind: "renderedPageBreak",
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name !== "sdt") {
      leadingRenderedPageBreakPending = false;
    }
    if (child.isText && child.text) {
      const formatting = extractRunFormatting(child.marks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({ formatting, marks: child.marks, paraDefaults });
      if (inTocParagraph) {
        stripTocHyperlinkStyle(formatting);
      }
      const run: TextRun = {
        kind: "text",
        text: child.text,
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
      return;
    }
    if (child.type.name === "symbol") {
      const attrs = expectSymbolAttrs(child);
      const text = decodeOoxmlSymbolCharacter(attrs.char);
      if (text === null) {
        return;
      }
      const formatting = extractRunFormatting(child.marks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({ formatting, marks: child.marks, paraDefaults });
      if (inTocParagraph) {
        stripTocHyperlinkStyle(formatting);
      }
      const alternateFontFamily = getFontAlternate(attrs.font, fontAlternates);
      formatting.fontFamily = attrs.font;
      if (alternateFontFamily) {
        formatting.alternateFontFamily = alternateFontFamily;
      }
      runs.push({
        kind: "text",
        text,
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "hardBreak") {
      runs.push({
        kind: "lineBreak",
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "tab") {
      const formatting = extractRunFormatting(child.marks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({ formatting, marks: child.marks, paraDefaults });
      const run: TabRun = {
        kind: "tab",
        ...mergeRunFormatting(paraDefaults, formatting),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
      return;
    }
    if (child.type.name === "image") {
      const attrs = expectImageAttrs(child);
      if (!attrs.src) {
        // Unsupported DrawingML shapes can survive the parser as image nodes
        // without a relationship target. They have no paintable payload; a
        // 100x100 fallback box would render a broken image and incorrectly
        // consume paragraph flow. Keep the host paragraph, but omit the run.
        return;
      }
      const constrained = constrainImageToPage(
        attrs.width ?? 100,
        attrs.height ?? 100,
        _options.pageContentHeight,
      );
      // Lift tracked-change marks off the image node so an inserted/deleted
      // picture paints in the revision colour and resolves with the rest of
      // the change. eigenpal #641.
      const trackedFmt = extractRunFormatting(child.marks, theme, fontAlternates);
      const run = buildImageRun(
        attrs,
        constrained,
        childPos,
        childPos + child.nodeSize,
        trackedFmt,
      );
      runs.push(run);
      return;
    }
    if (child.type.name === "field" || child.type.name === "structuredField") {
      // Marks on the field node (bold/italic/underline applied to the
      // field result inside `<w:fldChar separate>...</w:fldChar end>`)
      // must propagate to the run formatting, otherwise complex REF
      // fields whose visible text was authored as underlined (e.g.
      // cross-references like "Exhibit A" / "Section 1.3" in NVCA-style
      // templates) render with no underline. Reuse the same extractor
      // text runs use.
      const attrs = expectFieldAttrs(child);
      const ft = attrs.fieldType;
      let mappedType: FieldRun["fieldType"] = "OTHER";
      if (ft === "PAGE") {
        mappedType = "PAGE";
      } else if (ft === "NUMPAGES") {
        mappedType = "NUMPAGES";
      } else if (ft === "DATE") {
        mappedType = "DATE";
      } else if (ft === "TIME") {
        mappedType = "TIME";
      }
      const extractedFieldFormatting = extractRunFormatting(child.marks, theme, fontAlternates);
      applyCharacterStyleToggleFormatting({
        formatting: extractedFieldFormatting,
        marks: child.marks,
        paraDefaults,
      });
      if (inTocParagraph) {
        stripTocHyperlinkStyle(extractedFieldFormatting);
      }
      const fieldFormatting = markDefaultBlackTextColorSource(
        extractedFieldFormatting,
        paraDefaults,
      );
      const run: FieldRun = {
        kind: "field",
        fieldType: mappedType,
        instruction: attrs.instruction,
        fallback: _options.numberedRefResults?.get(child) ?? attrs.displayText ?? "",
        ...(attrs.fldLock ? { fldLock: true } : {}),
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
        ...fieldFormatting,
      };
      runs.push(run);
      return;
    }
    if (child.type.name === "math") {
      const attrs = expectMathAttrs(child);
      const plainText = attrs.plainText || "[equation]";
      runs.push({
        kind: "math",
        display: attrs.display ?? "inline",
        ommlXml: attrs.ommlXml,
        plainText,
        italic: true,
        fontFamily: "Cambria Math",
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      });
      return;
    }
    if (child.type.name === "sdt") {
      const sdtInnerOffset = childPos + 1; // +1 for opening tag
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
      child.forEach((sdtChild, sdtChildOffset) => {
        pushRunsForChild(sdtChild, sdtInnerOffset + sdtChildOffset);
      });
    }
  }

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, childOffset) => {
    pushRunsForChild(child, offset + childOffset);
  });

  return runs;
}

/**
 * Convert PM paragraph attrs to layout engine paragraph attrs.
 */
type ListMarkerRevision = NonNullable<ParagraphAttrs["listMarkerRevision"]>;

type ListPropertyChange = ParagraphPropertyChangeAttrs;
type ListPropertyFormatting = NonNullable<ListPropertyChange["previousFormatting"]>;

function toListMarkerRevision(
  kind: ListMarkerRevision["kind"],
  info: ListPropertyChange["info"],
): ListMarkerRevision {
  const revision: ListMarkerRevision = { kind };
  if (typeof info?.author === "string") {
    revision.author = info.author;
  }
  if (typeof info?.date === "string") {
    revision.date = info.date;
  }
  if (typeof info?.id === "number") {
    revision.revisionId = info.id;
  }
  return revision;
}

function isAddedNumberingChange(
  change: ListPropertyChange,
): change is ListPropertyChange & { previousFormatting: ListPropertyFormatting } {
  const previousFormatting = change.previousFormatting;
  return (
    previousFormatting != null &&
    Object.hasOwn(previousFormatting, "numPr") &&
    previousFormatting.numPr == null
  );
}

function isRemovedNumberingChange(
  change: ListPropertyChange,
): change is ListPropertyChange & { previousFormatting: ListPropertyFormatting } {
  const previousFormatting = change.previousFormatting;
  return (
    previousFormatting != null &&
    Object.hasOwn(previousFormatting, "numPr") &&
    isListNumPr(previousFormatting.numPr)
  );
}

function isChangedNumberingChange(
  currentNumPr: NonNullable<PMParagraphAttrs["numPr"]>,
  change: ListPropertyChange,
): change is ListPropertyChange & { previousFormatting: ListPropertyFormatting } {
  const previousFormatting = change.previousFormatting;
  return (
    previousFormatting != null &&
    Object.hasOwn(previousFormatting, "numPr") &&
    isListNumPr(previousFormatting.numPr) &&
    !areListNumPrEqual(previousFormatting.numPr, currentNumPr)
  );
}

function areListNumPrEqual(
  left: NonNullable<PMParagraphAttrs["numPr"]>,
  right: NonNullable<PMParagraphAttrs["numPr"]>,
): boolean {
  return left.numId === right.numId && left.ilvl === right.ilvl;
}

function isListNumPr(
  value: PMParagraphAttrs["numPr"] | null | undefined,
): value is NonNullable<PMParagraphAttrs["numPr"]> {
  return value !== undefined && value !== null;
}

function toPreviousListAttrs(previousFormatting: ListPropertyFormatting): PMParagraphAttrs {
  const attrs: PMParagraphAttrs = {};
  const numPr = previousFormatting.numPr;
  if (isListNumPr(numPr)) {
    attrs.numPr = numPr;
  }

  const listIsBullet = previousFormatting.listIsBullet;
  if (listIsBullet !== undefined) {
    attrs.listIsBullet = listIsBullet;
  }

  const listIsLegal = previousFormatting.listIsLegal;
  if (listIsLegal !== undefined) {
    attrs.listIsLegal = listIsLegal;
  }

  const listMarker = previousFormatting.listMarker;
  if (listMarker !== undefined) {
    attrs.listMarker = listMarker;
  }

  const listNumFmt = previousFormatting.listNumFmt;
  if (listNumFmt !== undefined) {
    attrs.listNumFmt = listNumFmt;
  }

  const listLevelNumFmts = previousFormatting.listLevelNumFmts;
  if (listLevelNumFmts !== undefined) {
    attrs.listLevelNumFmts = listLevelNumFmts;
  }

  const listLevelStarts = previousFormatting.listLevelStarts;
  if (listLevelStarts !== undefined) {
    attrs.listLevelStarts = listLevelStarts;
  }

  const listAbstractNumId = previousFormatting.listAbstractNumId;
  if (listAbstractNumId !== undefined) {
    attrs.listAbstractNumId = listAbstractNumId;
  }

  const listStartOverride = previousFormatting.listStartOverride;
  if (listStartOverride !== undefined) {
    attrs.listStartOverride = listStartOverride;
  }

  const listMarkerHidden = previousFormatting.listMarkerHidden;
  if (listMarkerHidden !== undefined) {
    attrs.listMarkerHidden = listMarkerHidden;
  }

  const listMarkerFormatting = previousFormatting.listMarkerFormatting;
  if (listMarkerFormatting !== undefined) {
    attrs.listMarkerFormatting = listMarkerFormatting;
  }

  const listMarkerAlignment = previousFormatting.listMarkerAlignment;
  if (listMarkerAlignment !== undefined) {
    attrs.listMarkerAlignment = listMarkerAlignment;
  }

  const listMarkerSuffix = previousFormatting.listMarkerSuffix;
  if (listMarkerSuffix !== undefined) {
    attrs.listMarkerSuffix = listMarkerSuffix;
  }

  return attrs;
}

function resolveDeletedListMarker(
  previousListAttrs: PMParagraphAttrs,
  listCounterState: ListCounterState | undefined,
): string | null {
  if (listCounterState && previousListAttrs.numPr) {
    // Advance the original counter stream in place (no clone): a
    // removed-numbering deletion occupied a number in the pre-revision
    // document, so it must progress the counter exactly like a deleted list
    // item — otherwise a following deletion on the same numId reuses it.
    const marker = computeListMarker(previousListAttrs, listCounterState);
    if (marker) {
      return marker;
    }
  }

  if (previousListAttrs.listMarker) {
    return previousListAttrs.listIsBullet
      ? convertBulletToUnicode(previousListAttrs.listMarker)
      : previousListAttrs.listMarker;
  }

  if (previousListAttrs.listIsBullet) {
    return "\u2022";
  }

  return null;
}

function applyDeletedListMarkerAttrs(
  attrs: ParagraphAttrs,
  change: ListPropertyChange & { previousFormatting: ListPropertyFormatting },
  listCounterState: ListCounterState | undefined,
  theme: Theme | null | undefined,
  fontAlternates: FontAlternates | undefined,
): void {
  const previousListAttrs = toPreviousListAttrs(change.previousFormatting);
  const marker = resolveDeletedListMarker(previousListAttrs, listCounterState);
  if (!marker) {
    return;
  }

  attrs.listMarker = marker;
  attrs.listMarkerRevision = toListMarkerRevision("del", change.info);
  if (previousListAttrs.listIsBullet !== undefined) {
    attrs.listIsBullet = previousListAttrs.listIsBullet;
  }
  if (previousListAttrs.listMarkerHidden !== undefined) {
    attrs.listMarkerHidden = previousListAttrs.listMarkerHidden;
  }
  if (previousListAttrs.listMarkerFormatting) {
    attrs.listMarkerFormatting = textFormattingToRunFormatting(
      previousListAttrs.listMarkerFormatting,
      theme,
      fontAlternates,
    );
  }
  if (previousListAttrs.listMarkerAlignment) {
    attrs.listMarkerAlignment = previousListAttrs.listMarkerAlignment;
  }
  if (previousListAttrs.listMarkerSuffix) {
    attrs.listMarkerSuffix = previousListAttrs.listMarkerSuffix;
  }
}

type ConvertParagraphAttrsOptions = {
  theme: Theme | null | undefined;
  fontAlternates: FontAlternates | undefined;
  listCounterStreams: ListCounterStreams;
  defaultTabStopTwips: number | undefined;
};

function convertParagraphAttrs(
  pmAttrs: PMParagraphAttrs,
  { theme, fontAlternates, listCounterStreams, defaultTabStopTwips }: ConvertParagraphAttrsOptions,
): ParagraphAttrs {
  const attrs: ParagraphAttrs = {};

  // Alignment - map DOCX values to CSS-compatible values
  // DOCX uses 'both' for justify, 'distribute' for distributed justify
  if (pmAttrs.alignment) {
    const align = pmAttrs.alignment;
    if (align === "both" || align === "distribute") {
      attrs.alignment = "justify";
    } else if (align === "left") {
      attrs.alignment = "left";
    } else if (align === "center") {
      attrs.alignment = "center";
    } else if (align === "right") {
      attrs.alignment = "right";
    }
    // Other DOCX alignments (mediumKashida, highKashida, lowKashida, thaiDistribute, justify)
    // default to no alignment set (inherits from style or defaults to left)
  }

  if (typeof pmAttrs.outlineLevel === "number") {
    attrs.outlineLevel = pmAttrs.outlineLevel;
  }

  // Spacing. HTML-origin auto spacing (w:beforeAutospacing/afterAutospacing)
  // renders Word's 14pt auto gap, overriding the imported before/after (which
  // Word writes as `0`) — surface it here so pagination matches the rendered
  // margins (eigenpal/docx-editor#823). But only while the effective spacing
  // still matches the import baseline; any later command or style change that
  // writes a different spacing value must win over the stale auto-spacing flag.
  const spaceBefore = pmAttrs.spaceBefore;
  const spaceAfter = pmAttrs.spaceAfter;
  const lineSpacing = pmAttrs.lineSpacing;
  if (typeof pmAttrs.snapToGrid === "boolean") {
    attrs.snapToGrid = pmAttrs.snapToGrid;
  }
  const autoBefore = autospacingMatchesBase(pmAttrs._autospacingBase, "before", spaceBefore);
  const autoAfter = autospacingMatchesBase(pmAttrs._autospacingBase, "after", spaceAfter);
  if (
    autoBefore ||
    autoAfter ||
    typeof spaceBefore === "number" ||
    typeof spaceAfter === "number" ||
    typeof lineSpacing === "number"
  ) {
    attrs.spacing = {};
    if (autoBefore) {
      attrs.spacing.before = AUTO_PARAGRAPH_SPACING_PX;
    } else if (typeof spaceBefore === "number") {
      attrs.spacing.before = twipsToPixels(spaceBefore);
    }
    if (autoAfter) {
      attrs.spacing.after = AUTO_PARAGRAPH_SPACING_PX;
    } else if (typeof spaceAfter === "number") {
      attrs.spacing.after = twipsToPixels(spaceAfter);
    }
    if (autoBefore || autoAfter) {
      attrs.automaticSpacing = {
        ...(autoBefore ? { before: true } : {}),
        ...(autoAfter ? { after: true } : {}),
      };
    }
    // Preserve spacing sides whose source Word renders on an empty paragraph:
    // direct formatting, document defaults, the implicit default paragraph
    // style, and automatic spacing. Layout consumes the combined provenance;
    // the authored PM attributes remain source-specific for serialization.
    const pmSpacingExplicit = pmAttrs.spacingExplicit as
      | { before?: boolean; after?: boolean }
      | null
      | undefined;
    const spacingFromDocDefaults = pmAttrs.spacingFromDocDefaults as
      | { before?: boolean; after?: boolean }
      | null
      | undefined;
    const spacingFromImplicitDefaultStyle = pmAttrs.spacingFromImplicitDefaultStyle as
      | { before?: boolean; after?: boolean }
      | null
      | undefined;
    const explicit: { before?: boolean; after?: boolean } = {};
    if (
      autoBefore ||
      pmSpacingExplicit?.before ||
      spacingFromDocDefaults?.before ||
      spacingFromImplicitDefaultStyle?.before
    ) {
      explicit.before = true;
    }
    if (
      autoAfter ||
      pmSpacingExplicit?.after ||
      spacingFromDocDefaults?.after ||
      spacingFromImplicitDefaultStyle?.after
    ) {
      explicit.after = true;
    }
    if (explicit.before !== undefined || explicit.after !== undefined) {
      attrs.spacingExplicit = explicit;
    }
    if (typeof lineSpacing === "number") {
      // Line spacing in twips - convert to multiplier or exact
      if (pmAttrs.lineSpacingRule === "exact" || pmAttrs.lineSpacingRule === "atLeast") {
        attrs.spacing.line = twipsToPixels(lineSpacing);
        attrs.spacing.lineUnit = "px";
        attrs.spacing.lineRule = pmAttrs.lineSpacingRule;
      } else {
        // Auto - line spacing is in 240ths of a line
        attrs.spacing.line = lineSpacing / 240;
        attrs.spacing.lineUnit = "multiplier";
        attrs.spacing.lineRule = "auto";
      }
    }
  }

  // Indentation - handle list item fallback calculation
  // For list items without explicit indentation, calculate based on level
  let indentLeft = typeof pmAttrs.indentLeft === "number" ? pmAttrs.indentLeft : undefined;
  let indentFirstLine =
    typeof pmAttrs.indentFirstLine === "number" ? pmAttrs.indentFirstLine : undefined;
  let hangingIndent = pmAttrs.hangingIndent;
  if (pmAttrs.numPr?.numId && indentLeft === undefined && indentFirstLine === undefined) {
    // Fallback: calculate indentation based on level
    // An authored first-line or hanging position is already a complete list
    // marker anchor. Adding a synthetic left indent would shift that anchor a
    // second time, while tab stops still resolve from the paragraph margin.
    // Each level indents 0.5 inch (720 twips) more
    const level = pmAttrs.numPr.ilvl ?? 0;
    // Base indentation: 0.5 inch (720 twips) per level
    // Level 0 = 720 twips, Level 1 = 1440 twips, etc.
    indentLeft = (level + 1) * 720;
    // Default hanging indent of 360 twips for the list marker
    if (indentFirstLine === undefined) {
      indentFirstLine = -360;
      hangingIndent = true;
    }
  }

  if (
    indentLeft !== undefined ||
    typeof pmAttrs.indentRight === "number" ||
    indentFirstLine !== undefined
  ) {
    attrs.indent = {};
    if (indentLeft !== undefined) {
      attrs.indent.left = twipsToPixels(indentLeft);
    }
    if (typeof pmAttrs.indentRight === "number") {
      attrs.indent.right = twipsToPixels(pmAttrs.indentRight);
    }
    if (indentFirstLine !== undefined) {
      if (hangingIndent) {
        // Hanging indent: indentFirstLine is stored as negative, convert to positive for rendering
        attrs.indent.hanging = Math.abs(twipsToPixels(indentFirstLine));
      } else {
        attrs.indent.firstLine = twipsToPixels(indentFirstLine);
      }
    }
  }

  // Style ID
  if (pmAttrs.styleId) {
    attrs.styleId = pmAttrs.styleId;
  }

  // Borders
  if (pmAttrs.borders) {
    const borders = pmAttrs.borders;
    attrs.borders = {};

    const convertBorder = (border: typeof borders.top) =>
      border ? convertBorderSpecToLayout(border, theme) : undefined;

    const topBorder = borders.top ? convertBorder(borders.top) : undefined;
    if (topBorder) {
      attrs.borders.top = topBorder;
    }
    const bottomBorder = borders.bottom ? convertBorder(borders.bottom) : undefined;
    if (bottomBorder) {
      attrs.borders.bottom = bottomBorder;
    }
    const leftBorder = borders.left ? convertBorder(borders.left) : undefined;
    if (leftBorder) {
      attrs.borders.left = leftBorder;
    }
    const rightBorder = borders.right ? convertBorder(borders.right) : undefined;
    if (rightBorder) {
      attrs.borders.right = rightBorder;
    }
    const betweenBorder = borders.between ? convertBorder(borders.between) : undefined;
    if (betweenBorder) {
      attrs.borders.between = betweenBorder;
    }
    const barBorder = borders.bar ? convertBorder(borders.bar) : undefined;
    if (barBorder) {
      attrs.borders.bar = barBorder;
    }

    // Only include if at least one border is set
    if (
      !attrs.borders.top &&
      !attrs.borders.bottom &&
      !attrs.borders.left &&
      !attrs.borders.right &&
      !attrs.borders.between &&
      !attrs.borders.bar
    ) {
      delete attrs.borders;
    }
  }

  // Shading (background color). Word's `Normal` paragraph style commonly
  // sets `<w:shd val="clear" fill="FFFFFF"/>` — semantically a no-op on
  // a white page, but folio's dark mode draws the literal `#FFFFFF`
  // fill as a visible white block over the dark canvas. Treat any white
  // shading as transparent (= page background) so it renders the same as
  // "no shading" in both modes. Other shading colors are preserved
  // verbatim so authored highlights stay visible.
  const shadingRgb = pmAttrs.shading?.fill?.rgb?.toUpperCase();
  if (shadingRgb && shadingRgb !== "FFFFFF" && shadingRgb !== "FFFFFE") {
    attrs.shading = `#${pmAttrs.shading?.fill?.rgb}`;
  }

  // Tab stops
  if (pmAttrs.tabs && pmAttrs.tabs.length > 0) {
    attrs.tabs = pmAttrs.tabs.map((tab) => {
      const tabStop: TabStop = {
        val: mapTabAlignment(tab.alignment),
        pos: tab.position,
      };
      if (tab.leader) {
        tabStop.leader = tab.leader as NonNullable<TabStop["leader"]>;
      }
      return tabStop;
    });
  }

  // Page break control
  if (pmAttrs.pageBreakBefore) {
    attrs.pageBreakBefore = true;
  }
  if (pmAttrs.kinsoku !== undefined && pmAttrs.kinsoku !== null) {
    attrs.kinsoku = pmAttrs.kinsoku;
  }
  if (pmAttrs.overflowPunctuation !== undefined && pmAttrs.overflowPunctuation !== null) {
    attrs.overflowPunctuation = pmAttrs.overflowPunctuation;
  }
  if (pmAttrs.suppressAutoHyphens !== undefined && pmAttrs.suppressAutoHyphens !== null) {
    attrs.suppressAutoHyphens = pmAttrs.suppressAutoHyphens;
  }
  if (pmAttrs.renderedPageBreakBefore) {
    attrs.renderedPageBreakBefore = true;
  }
  if (pmAttrs.keepNext) {
    attrs.keepNext = true;
  }
  if (pmAttrs.keepLines) {
    attrs.keepLines = true;
  }
  if (pmAttrs.widowControl === false) {
    attrs.widowControl = false;
  }
  if (pmAttrs.contextualSpacing) {
    attrs.contextualSpacing = true;
  }
  if (pmAttrs.runInWithNext) {
    attrs.runInWithNext = true;
  }
  const bidi = directionToBidi(pmAttrs.direction);
  if (bidi !== undefined) {
    attrs.bidi = bidi;
  }
  if (pmAttrs.styleId) {
    attrs.styleId = pmAttrs.styleId;
  }

  // List properties
  const propertyChanges = pmAttrs._propertyChanges ?? [];
  let changedNumberingChange:
    | (ListPropertyChange & { previousFormatting: ListPropertyFormatting })
    | undefined;
  if (pmAttrs.numPr) {
    const numPr: ParagraphAttrs["numPr"] & object = {};
    if (pmAttrs.numPr.numId !== undefined) {
      numPr.numId = pmAttrs.numPr.numId;
    }
    if (pmAttrs.numPr.ilvl !== undefined) {
      numPr.ilvl = pmAttrs.numPr.ilvl;
    }
    attrs.numPr = numPr;

    if (pmAttrs.pPrMark?.kind === "del") {
      attrs.listMarkerRevision = toListMarkerRevision("del", pmAttrs.pPrMark.info);
    } else if (pmAttrs.pPrMark?.kind === "ins") {
      attrs.listMarkerRevision = toListMarkerRevision("ins", pmAttrs.pPrMark.info);
    } else {
      const numberingAddedChange = propertyChanges.find(isAddedNumberingChange);
      const currentNumPr = pmAttrs.numPr;
      changedNumberingChange = propertyChanges.find(
        (
          change,
        ): change is ListPropertyChange & {
          previousFormatting: ListPropertyFormatting;
        } => isChangedNumberingChange(currentNumPr, change),
      );
      const numberingInsertionChange = numberingAddedChange ?? changedNumberingChange;
      if (numberingInsertionChange) {
        attrs.listMarkerRevision = toListMarkerRevision("ins", numberingInsertionChange.info);
      }
    }
  }
  const visibleMarker = advanceVisibleListMarker(pmAttrs, listCounterStreams);
  const resolvedMarker = applyMarkerAllCaps(visibleMarker.marker, pmAttrs.listMarkerAllCaps);
  if (resolvedMarker !== null) {
    attrs.listMarker = resolvedMarker;
  } else if (pmAttrs.listMarker) {
    attrs.listMarker = pmAttrs.listIsBullet
      ? convertBulletToUnicode(pmAttrs.listMarker)
      : pmAttrs.listMarker;
  }
  if (pmAttrs.listIsBullet !== undefined) {
    attrs.listIsBullet = pmAttrs.listIsBullet;
  }
  if (pmAttrs.listMarkerHidden) {
    attrs.listMarkerHidden = true;
  }
  if (pmAttrs.listMarkerFormatting) {
    attrs.listMarkerFormatting = textFormattingToRunFormatting(
      pmAttrs.listMarkerFormatting,
      theme,
      fontAlternates,
    );
  }
  if (pmAttrs.listMarkerAlignment) {
    attrs.listMarkerAlignment = pmAttrs.listMarkerAlignment;
  }
  if (pmAttrs.listMarkerSuffix) {
    attrs.listMarkerSuffix = pmAttrs.listMarkerSuffix;
  }
  if (pmAttrs.listMarkerSecondSlotOffsetTwips !== undefined) {
    attrs.listMarkerSecondSlotOffsetTwips = pmAttrs.listMarkerSecondSlotOffsetTwips;
  }
  if (!pmAttrs.numPr) {
    const numberingRemovedChange = propertyChanges.find(isRemovedNumberingChange);
    if (numberingRemovedChange) {
      // Number removed-numbering deletions off the original stream too (like
      // deleted list items): the struck-through marker must reflect the
      // pre-revision number, not the final counter that insertions advanced.
      applyDeletedListMarkerAttrs(attrs, numberingRemovedChange, undefined, theme, fontAlternates);
      if (resolvedMarker !== null) {
        attrs.listMarker = resolvedMarker;
        attrs.listMarkerRevision = toListMarkerRevision("del", numberingRemovedChange.info);
      }
    }
  }
  if (defaultTabStopTwips !== undefined) {
    attrs.defaultTabStopTwips = defaultTabStopTwips;
  }
  // Default font for empty paragraph measurement (from style's rPr / pPr/rPr)
  const dtf = pmAttrs.defaultTextFormatting as TextFormatting | undefined;
  if (dtf) {
    if (dtf.fontSize !== undefined) {
      // fontSize in TextFormatting is in half-points, convert to points
      attrs.defaultFontSize = dtf.fontSize / 2;
    }
    if (
      attrs.listMarker !== undefined &&
      !attrs.listMarkerHidden &&
      dtf.fontSizeCs !== undefined &&
      (dtf.fontSize === undefined || dtf.fontSizeCs > dtf.fontSize)
    ) {
      attrs.listParagraphMarkFontSize = dtf.fontSizeCs / 2;
    }
    if (dtf.fontFamily) {
      const resolvedFamily = resolveWesternThemeFont(dtf.fontFamily, theme);
      if (resolvedFamily) {
        attrs.defaultFontFamily = resolvedFamily;
        const alternate = getFontAlternate(resolvedFamily, fontAlternates);
        if (alternate) {
          attrs.defaultAlternateFontFamily = alternate;
        }
      }
    }
  }

  return attrs;
}

/**
 * Map document TabStopAlignment to layout engine TabAlignment
 */
function mapTabAlignment(
  align: "left" | "center" | "right" | "decimal" | "bar" | "clear" | "num",
): "start" | "end" | "center" | "decimal" | "bar" | "clear" {
  switch (align) {
    case "left":
      return "start";
    case "right":
      return "end";
    case "center":
      return "center";
    case "decimal":
      return "decimal";
    case "bar":
      return "bar";
    case "clear":
      return "clear";
    case "num":
      return "start"; // Number tab treated as left-aligned
    default:
      return "start";
  }
}

/**
 * Convert a paragraph node to a ParagraphBlock.
 */
function hasOnlyVisuallyEmptyTextRuns(runs: Run[]): boolean {
  return (
    runs.length > 0 &&
    runs.every(
      (run) => run.kind === "text" && run.text.replace(/\u00a0/gu, " ").trim().length === 0,
    )
  );
}

function convertParagraph(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
): ParagraphBlock {
  const pmAttrs = expectParagraphAttrs(node);
  const runs = paragraphToRuns(node, startPos, options);
  const attrs = convertParagraphAttrs(pmAttrs, {
    theme: options.theme,
    fontAlternates: options.fontAlternates,
    listCounterStreams: options.listCounterStreams,
    defaultTabStopTwips: options.defaultTabStopTwips,
  });
  if (options.lineBreakRules) {
    attrs.lineBreakRules = options.lineBreakRules;
  }
  if (options.justificationCompatibility) {
    attrs.justificationCompatibility = options.justificationCompatibility;
  }
  if (options.automaticHyphenation) {
    attrs.automaticHyphenation = options.automaticHyphenation;
  }
  const defaultTextFormatting = pmAttrs.defaultTextFormatting as TextFormatting | undefined;
  if (runs.length === 0 || hasOnlyVisuallyEmptyTextRuns(runs)) {
    const hasDirectParagraphFormatting =
      pmAttrs._originalFormatting &&
      Object.entries(pmAttrs._originalFormatting).some(
        ([key, value]) => key !== "runProperties" && value !== undefined && value !== null,
      );
    if (hasDirectParagraphFormatting) {
      attrs.hasDirectParagraphFormatting = true;
    }
    const paragraphMarkFormatting = pmAttrs._originalFormatting?.runProperties;
    if (
      paragraphMarkFormatting &&
      Object.values(paragraphMarkFormatting).some((value) => value !== undefined && value !== null)
    ) {
      attrs.hasDirectParagraphMarkFormatting = true;
    }
    if (paragraphMarkFormatting?.fontSize !== undefined) {
      attrs.defaultFontSize = paragraphMarkFormatting.fontSize / 2;
    }
    const paragraphMarkFontFamily = paragraphMarkFormatting?.fontFamily
      ? resolveWesternThemeFont(paragraphMarkFormatting.fontFamily, options.theme)
      : undefined;
    if (paragraphMarkFontFamily) {
      attrs.defaultFontFamily = paragraphMarkFontFamily;
      const alternate = getFontAlternate(paragraphMarkFontFamily, options.fontAlternates);
      if (alternate) {
        attrs.defaultAlternateFontFamily = alternate;
      }
    }
  }
  if (runs.length === 0 && defaultTextFormatting?.hidden === true) {
    attrs.suppressEmptyParagraphHeight = true;
    if (attrs.listMarker !== undefined) {
      attrs.listMarkerHidden = true;
    }
  }
  const hasVisibleParagraphPayload =
    (attrs.listMarker !== undefined && !attrs.listMarkerHidden) ||
    attrs.borders?.top !== undefined ||
    attrs.borders?.bottom !== undefined ||
    attrs.borders?.left !== undefined ||
    attrs.borders?.right !== undefined ||
    attrs.borders?.between !== undefined ||
    attrs.borders?.bar !== undefined ||
    attrs.shading !== undefined;
  if (runs.length === 0 && pmAttrs._pageBreakCarrier === true && !hasVisibleParagraphPayload) {
    attrs.suppressEmptyParagraphHeight = true;
  }

  const bookmarkNames = pmAttrs.bookmarks?.map((b) => b.name);

  const block: ParagraphBlock = {
    kind: "paragraph",
    id: nextBlockId(),
    runs,
    attrs,
    ...(pmAttrs.paraId ? { paraId: pmAttrs.paraId } : {}),
    ...(bookmarkNames && bookmarkNames.length > 0 ? { bookmarks: bookmarkNames } : {}),
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  const frame = pmAttrs._originalFormatting?.frame;
  if (frame !== undefined && frame.dropCap !== "drop" && frame.dropCap !== "margin") {
    setParagraphFrame(block, {
      ...(frame.width !== undefined ? { width: twipsToPixels(frame.width) } : {}),
      ...(frame.height !== undefined ? { height: twipsToPixels(frame.height) } : {}),
      ...(frame.hSpace !== undefined ? { hSpace: twipsToPixels(frame.hSpace) } : {}),
      ...(frame.vSpace !== undefined ? { vSpace: twipsToPixels(frame.vSpace) } : {}),
      ...(frame.hAnchor !== undefined ? { hAnchor: frame.hAnchor } : {}),
      ...(frame.vAnchor !== undefined ? { vAnchor: frame.vAnchor } : {}),
      ...(frame.x !== undefined ? { x: twipsToPixels(frame.x) } : {}),
      ...(frame.y !== undefined ? { y: twipsToPixels(frame.y) } : {}),
      ...(frame.xAlign !== undefined ? { xAlign: frame.xAlign } : {}),
      ...(frame.yAlign !== undefined ? { yAlign: frame.yAlign } : {}),
      ...(frame.wrap !== undefined ? { wrap: frame.wrap } : {}),
    });
  }
  return block;
}

/**
 * Word keeps a final empty body paragraph after a table as an editable anchor,
 * but that final anchor does not create a page of its own. Earlier authored
 * empty paragraphs retain their height and may carry the document onto a blank
 * page. Preserve every block and PM range while collapsing only the final one.
 */
function isPaintlessTerminalParagraph(block: FlowBlock | undefined): block is ParagraphBlock {
  if (block?.kind !== "paragraph" || block.runs.length !== 0) {
    return false;
  }

  const attrs = block.attrs;
  return !(
    (attrs?.listMarker !== undefined && !attrs.listMarkerHidden) ||
    attrs?.borders?.top ||
    attrs?.borders?.bottom ||
    attrs?.borders?.left ||
    attrs?.borders?.right ||
    attrs?.borders?.between ||
    attrs?.borders?.bar ||
    attrs?.shading ||
    attrs?.spacingExplicit?.before ||
    attrs?.spacingExplicit?.after ||
    attrs?.pageBreakBefore ||
    attrs?.renderedPageBreakBefore
  );
}

function suppressFinalEmptyParagraphAfterTable(blocks: FlowBlock[]): void {
  let suffixStart = blocks.length;
  while (suffixStart > 0 && isPaintlessTerminalParagraph(blocks[suffixStart - 1])) {
    suffixStart -= 1;
  }

  if (
    suffixStart === blocks.length ||
    suffixStart === 0 ||
    blocks[suffixStart - 1]?.kind !== "table"
  ) {
    return;
  }

  const finalBlock = blocks.at(-1);
  if (isPaintlessTerminalParagraph(finalBlock)) {
    finalBlock.attrs = { ...finalBlock.attrs, suppressEmptyParagraphHeight: true };
  }
}

function suppressFinalParagraphInRepeatedEmptySuffix(blocks: FlowBlock[]): void {
  let suffixStart = blocks.length;
  while (suffixStart > 0 && isPaintlessTerminalParagraph(blocks[suffixStart - 1])) {
    suffixStart -= 1;
  }

  if (
    suffixStart === 0 ||
    blocks.length - suffixStart < 2 ||
    blocks[suffixStart - 1]?.kind === "table"
  ) {
    return;
  }

  const finalBlock = blocks.at(-1);
  if (isPaintlessTerminalParagraph(finalBlock)) {
    finalBlock.attrs = { ...finalBlock.attrs, suppressEmptyParagraphHeight: true };
  }
}

function reserveLeadingEmptyOutlineHeight(blocks: FlowBlock[]): void {
  const firstBlock = blocks.at(0);
  if (
    firstBlock?.kind !== "paragraph" ||
    firstBlock.runs.length !== 0 ||
    firstBlock.attrs?.outlineLevel !== 0
  ) {
    return;
  }

  firstBlock.attrs = { ...firstBlock.attrs, reserveEmptyOutlineHeight: true };
}

/**
 * Convert border width from eighths of a point to pixels.
 * OOXML stores border widths in eighths of a point.
 */
function borderWidthToPixels(eighthsOfPoint: number): number {
  return pointsToPixels(eighthsOfPoint / 8);
}

// OOXML border style → CSS border-style mapping
const OOXML_TO_CSS_BORDER: Record<string, string> = {
  single: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  thick: "solid",
  dashSmallGap: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
  triple: "double",
  wave: "solid",
  doubleWave: "double",
  threeDEmboss: "ridge",
  threeDEngrave: "groove",
  outset: "outset",
  inset: "inset",
};

/**
 * Convert an OOXML BorderSpec to a layout-engine BorderStyle.
 * Shared by paragraph borders, cell borders, and header/footer borders.
 */
export function convertBorderSpecToLayout(
  border: {
    style?: string;
    size?: number;
    space?: number;
    color?: {
      rgb?: string;
      themeColor?: string;
      themeTint?: string;
      themeShade?: string;
    };
  },
  theme?: Theme | null,
): BorderStyle | undefined {
  if (!border.style || border.style === "none" || border.style === "nil") {
    return undefined;
  }
  const result: BorderStyle = {
    style: OOXML_TO_CSS_BORDER[border.style] || "solid",
    width: border.size === undefined ? 1 : borderWidthToPixels(border.size),
    color: border.color
      ? resolveColor(border.color as Parameters<typeof resolveColor>[0], theme)
      : "#000000",
  };
  if (border.space !== undefined) {
    result.space = pointsToPixels(border.space);
  }
  return result;
}

/**
 * Extract cell borders from ProseMirror attributes.
 * Borders are full BorderSpec objects with style/size/color.
 */
function extractCellBorders(
  borders:
    | Record<
        string,
        {
          style?: string;
          size?: number;
          color?: {
            rgb?: string;
            themeColor?: string;
            themeTint?: string;
            themeShade?: string;
          };
        }
      >
    | null
    | undefined,
  theme?: Theme | null,
): CellBorders | undefined {
  if (!borders) {
    return undefined;
  }

  const result: CellBorders = {};
  const sides = ["top", "bottom", "left", "right"] as const;

  for (const side of sides) {
    const border = borders[side];
    const converted = border ? convertBorderSpecToLayout(border, theme) : undefined;
    if (!converted) {
      result[side] = { width: 0, style: "none" };
      continue;
    }

    result[side] = border?.size === 0 ? { ...converted, width: 0 } : converted;
  }

  const diagonalSides = ["topLeftToBottomRight", "topRightToBottomLeft"] as const;
  for (const side of diagonalSides) {
    const border = borders[side];
    const converted = border ? convertBorderSpecToLayout(border, theme) : undefined;
    if (converted) {
      result[side] = converted;
    }
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * Convert a table cell node.
 */
function convertTableCell(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
  tableCellMargins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  },
): TableCell {
  const blocks: FlowBlock[] = [];
  let offset = startPos + 1; // +1 for opening tag

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    if (child.type.name === "paragraph") {
      const block = convertParagraph(child, offset, options);
      blocks.push(block);
    } else if (child.type.name === "table") {
      blocks.push(convertTable(child, offset, options));
    } else if (child.type.name === "textBox") {
      blocks.push(convertTextBoxNode(child, offset, options));
    }
    offset += child.nodeSize;
  });

  // A table cell whose final block is a nested table needs a trailing paragraph
  // as its editable cell-end marker. Keep that marker as a zero-height anchor.
  // Empty paragraphs after ordinary prose are authored content and retain their
  // normal line height.
  const trailingBlock = blocks.at(-1);
  const precedingBlock = blocks.at(-2);
  if (
    precedingBlock?.kind === "table" &&
    trailingBlock?.kind === "paragraph" &&
    trailingBlock.runs.every((run) => run.kind === "text" && run.text.length === 0)
  ) {
    trailingBlock.attrs = { ...trailingBlock.attrs, suppressEmptyParagraphHeight: true };
  }

  const attrs = expectTableCellAttrs(node);
  if (
    attrs.hideMark &&
    trailingBlock?.kind === "paragraph" &&
    trailingBlock.runs.every((run) => run.kind === "text" && run.text.length === 0)
  ) {
    trailingBlock.attrs = { ...trailingBlock.attrs, suppressEmptyParagraphHeight: true };
  }

  // Convert cell margins (twips) to pixel padding
  // OOXML TableNormal defaults: top=0, bottom=0, left=108 twips (~7px), right=108 twips (~7px)
  const margins = attrs.margins;
  const resolvePaddingSide = (
    side: TablePaddingSide,
    cellTwips: number | undefined,
    tableTwips: number | undefined,
  ): number => {
    if (cellTwips !== undefined) {
      return twipsToPixels(cellTwips);
    }
    if (tableTwips !== undefined) {
      return twipsToPixels(tableTwips);
    }
    return twipsToPixels(DEFAULT_TABLE_CELL_MARGIN_TWIPS[side]);
  };
  const padding = {
    top: resolvePaddingSide("top", margins?.top, tableCellMargins?.top),
    right: resolvePaddingSide("right", margins?.right, tableCellMargins?.right),
    bottom: resolvePaddingSide("bottom", margins?.bottom, tableCellMargins?.bottom),
    left: resolvePaddingSide("left", margins?.left, tableCellMargins?.left),
  };

  const cell: TableCell = {
    id: nextBlockId(),
    blocks: groupParagraphFrames(blocks, nextBlockId),
    colSpan: attrs.colspan,
    rowSpan: attrs.rowspan,
    padding,
  };
  if (attrs.width) {
    cell.width = twipsToPixels(attrs.width);
  }
  if (attrs.verticalAlign) {
    cell.verticalAlign = attrs.verticalAlign;
  }
  if (attrs.textDirection) {
    cell.textDirection = attrs.textDirection;
  }
  if (attrs.backgroundColor) {
    cell.background = `#${attrs.backgroundColor}`;
  }
  const cellBorders = extractCellBorders(attrs.borders, options.theme);
  if (cellBorders) {
    cell.borders = cellBorders;
  }
  if (attrs.noWrap) {
    cell.noWrap = true;
  }
  return cell;
}

/**
 * Convert a table row node.
 */
function convertTableRow(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
  tableCellMargins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  },
): TableRow {
  const cells: TableCell[] = [];
  let offset = startPos + 1; // +1 for opening tag

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    if (child.type.name === "tableCell" || child.type.name === "tableHeader") {
      cells.push(convertTableCell(child, offset, options, tableCellMargins));
    }
    offset += child.nodeSize;
  });

  const attrs = expectTableRowAttrs(node);
  const row: TableRow = {
    id: nextBlockId(),
    cells,
  };
  if (attrs._originalFormatting?.gridBefore) {
    row.gridBefore = attrs._originalFormatting.gridBefore;
  }
  if (attrs._originalFormatting?.gridAfter) {
    row.gridAfter = attrs._originalFormatting.gridAfter;
  }
  if (attrs.height) {
    row.height = twipsToPixels(attrs.height);
  }
  if (attrs.heightRule) {
    row.heightRule = attrs.heightRule;
  }
  if (attrs.isHeader) {
    row.isHeader = attrs.isHeader;
  }
  if (attrs._originalFormatting?.cantSplit) {
    row.cantSplit = true;
  }
  if (attrs.hidden) {
    row.hidden = attrs.hidden;
  }
  const effectiveJustification =
    attrs._originalFormatting?.justification ?? attrs._resolvedJustification;
  if (effectiveJustification) {
    row.justification = effectiveJustification;
  }
  return row;
}

/**
 * Convert a table node to a TableBlock.
 */
function convertTable(node: PMNode, startPos: number, options: FlowConversionOptions): TableBlock {
  const rows: TableRow[] = [];
  let offset = startPos + 1; // +1 for opening tag
  const attrs = expectTableAttrs(node);
  const tableCellMargins = attrs.cellMargins;

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    if (child.type.name === "tableRow") {
      rows.push(convertTableRow(child, offset, options, tableCellMargins));
    }
    offset += child.nodeSize;
  });

  // Extract columnWidths from node attributes and convert from twips to pixels
  const columnWidthsTwips = attrs.columnWidths;
  let columnWidths = columnWidthsTwips?.map(twipsToPixels);

  const width = attrs.width;
  const widthType = attrs.widthType;

  // Fallback: compute column widths from first row cell widths if table attr is missing
  if (!columnWidths && rows.length > 0) {
    // SAFETY: rows.length > 0 verified by condition above
    const firstRow = rows[0]!;
    const cellWidths = firstRow.cells.map((cell) => cell.width);
    // Only use if all cells have widths defined
    if (cellWidths.every((w) => w !== undefined && w > 0)) {
      columnWidths = cellWidths as number[];
    }
  }

  // Keep authored justification separate in ProseMirror so style-derived
  // placement never becomes direct formatting on save.
  const justification = attrs.justification ?? attrs._resolvedJustification;

  // Extract table indent + RTL column order from _originalFormatting
  // (w:tblInd, w:bidiVisual). bidiVisual is import-only — folio has no UI to
  // toggle it — so reading the preserved formatting is sufficient
  // (eigenpal/docx-editor#940).
  const originalFormatting = attrs._originalFormatting;
  const resolvedIndent = attrs._resolvedIndent;
  const effectiveIndent = resolvedIndent ?? originalFormatting?.indent;
  const indentPx =
    effectiveIndent?.value !== undefined && effectiveIndent?.type === "dxa"
      ? twipsToPixels(effectiveIndent.value)
      : undefined;

  const floating = attrs.floating as
    | {
        horzAnchor?: "margin" | "page" | "text";
        vertAnchor?: "margin" | "page" | "text";
        tblpX?: number;
        tblpXSpec?: "left" | "center" | "right" | "inside" | "outside";
        tblpY?: number;
        tblpYSpec?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
        topFromText?: number;
        bottomFromText?: number;
        leftFromText?: number;
        rightFromText?: number;
      }
    | undefined;

  let floatingPx: FloatingTablePosition | undefined;
  if (floating) {
    const fp: FloatingTablePosition = {};
    if (floating.horzAnchor) {
      fp.horzAnchor = floating.horzAnchor;
    }
    if (floating.vertAnchor) {
      fp.vertAnchor = floating.vertAnchor;
    }
    if (floating.tblpX !== undefined) {
      fp.tblpX = twipsToPixels(floating.tblpX);
    }
    if (floating.tblpXSpec) {
      fp.tblpXSpec = floating.tblpXSpec;
    }
    if (floating.tblpY !== undefined) {
      fp.tblpY = twipsToPixels(floating.tblpY);
    }
    if (floating.tblpYSpec) {
      fp.tblpYSpec = floating.tblpYSpec;
    }
    if (floating.topFromText !== undefined) {
      fp.topFromText = twipsToPixels(floating.topFromText);
    }
    if (floating.bottomFromText !== undefined) {
      fp.bottomFromText = twipsToPixels(floating.bottomFromText);
    }
    if (floating.leftFromText !== undefined) {
      fp.leftFromText = twipsToPixels(floating.leftFromText);
    }
    if (floating.rightFromText !== undefined) {
      fp.rightFromText = twipsToPixels(floating.rightFromText);
    }
    floatingPx = fp;
  }

  const tableBlock: TableBlock = {
    kind: "table",
    id: nextBlockId(),
    rows,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  if (columnWidths) {
    tableBlock.columnWidths = columnWidths;
  }
  if (width !== undefined) {
    tableBlock.width = width;
  }
  if (widthType !== undefined) {
    tableBlock.widthType = widthType;
  }
  if (originalFormatting?.layout !== undefined) {
    tableBlock.layout = originalFormatting.layout;
  }
  if (justification) {
    tableBlock.justification = justification;
  }
  if (indentPx !== undefined) {
    tableBlock.indent = indentPx;
  }
  if (floatingPx) {
    tableBlock.floating = floatingPx;
  }
  const effectiveBidi = attrs._resolvedBidi ?? originalFormatting?.bidi;
  if (effectiveBidi) {
    tableBlock.bidi = true;
  }
  return tableBlock;
}

/**
 * Convert an image node to an ImageBlock.
 */
function convertImage(node: PMNode, startPos: number, pageContentHeight?: number): ImageBlock {
  const attrs = expectImageAttrs(node);
  const wrapType = attrs.wrapType;

  // Only anchor images with 'behind' or 'inFront' wrap types
  // Other wrap types (square, tight, through, topAndBottom) need text wrapping
  // which we don't support yet, so treat them as block-level images
  const shouldAnchor = wrapType === "behind" || wrapType === "inFront";

  const constrained = constrainImageToPage(
    attrs.width ?? 100,
    attrs.height ?? 100,
    pageContentHeight,
  );

  const imgBlock: ImageBlock = {
    kind: "image",
    id: nextBlockId(),
    src: attrs.src,
    width: constrained.width,
    height: constrained.height,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  if (attrs.alt) {
    imgBlock.alt = attrs.alt;
  }
  if (attrs.transform) {
    imgBlock.transform = attrs.transform;
  }
  // eigenpal #424 (opacity render pipeline). `!= null` so PM's null schema
  // default doesn't leak into ImageBlock.opacity (`number | undefined`).
  if (attrs.opacity != null) {
    imgBlock.opacity = attrs.opacity;
  }
  if (shouldAnchor) {
    const anchor: NonNullable<ImageBlock["anchor"]> = {
      isAnchored: true,
      behindDoc: wrapType === "behind",
    };
    if (attrs.distLeft !== undefined) {
      anchor.offsetH = attrs.distLeft;
    }
    if (attrs.distTop !== undefined) {
      anchor.offsetV = attrs.distTop;
    }
    imgBlock.anchor = anchor;
  }
  if (attrs.hlinkHref) {
    imgBlock.hlinkHref = attrs.hlinkHref;
  }
  // eigenpal #424: thread wp:srcRect crop fractions to the floating-image
  // block so renderers can apply clip-path consistently across paths.
  if (attrs.cropTop != null) {
    imgBlock.cropTop = attrs.cropTop;
  }
  if (attrs.cropRight != null) {
    imgBlock.cropRight = attrs.cropRight;
  }
  if (attrs.cropBottom != null) {
    imgBlock.cropBottom = attrs.cropBottom;
  }
  if (attrs.cropLeft != null) {
    imgBlock.cropLeft = attrs.cropLeft;
  }
  // eigenpal #1096: preserve image border attrs for floating/block image
  // painting. PM defaults are null; treat null as absent.
  if (attrs.borderWidth != null) {
    imgBlock.borderWidth = attrs.borderWidth;
  }
  if (attrs.borderColor) {
    imgBlock.borderColor = attrs.borderColor;
  }
  if (attrs.borderStyle) {
    imgBlock.borderStyle = attrs.borderStyle;
  }
  return imgBlock;
}

/**
 * Convert a textBox PM node to a TextBoxBlock.
 */
function convertTextBoxNode(
  node: PMNode,
  startPos: number,
  opts: FlowConversionOptions,
): TextBoxBlock {
  const attrs = expectTextBoxAttrs(node);
  const contentBlocks: (ParagraphBlock | TableBlock)[] = [];

  // Convert child blocks inside the text box
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, offset) => {
    const childPos = startPos + 1 + offset;
    if (child.type.name === "paragraph") {
      contentBlocks.push(convertParagraph(child, childPos, opts));
      return;
    }
    if (child.type.name === "table") {
      contentBlocks.push(convertTable(child, childPos, opts));
    }
  });

  const textBox: TextBoxBlock = {
    kind: "textBox",
    id: nextBlockId(),
    width: attrs.width ?? DEFAULT_TEXTBOX_WIDTH,
    margins: {
      top: attrs.marginTop ?? DEFAULT_TEXTBOX_MARGINS.top,
      bottom: attrs.marginBottom ?? DEFAULT_TEXTBOX_MARGINS.bottom,
      left: attrs.marginLeft ?? DEFAULT_TEXTBOX_MARGINS.left,
      right: attrs.marginRight ?? DEFAULT_TEXTBOX_MARGINS.right,
    },
    content: contentBlocks,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  const verticalAlign = normalizeShapeTextAnchor(attrs.verticalAlign);
  if (verticalAlign !== undefined) {
    textBox.verticalAlign = verticalAlign;
  }
  if (attrs.height !== undefined) {
    textBox.height = attrs.height;
  }
  if (attrs.autoFit !== undefined) {
    textBox.autoFit = attrs.autoFit;
  }
  if (attrs.textWrap !== undefined) {
    textBox.textWrap = attrs.textWrap;
  }
  if (attrs.fillColor !== undefined) {
    textBox.fillColor = attrs.fillColor;
  }
  if (attrs.outlineWidth !== undefined) {
    textBox.outlineWidth = attrs.outlineWidth;
  }
  if (attrs.outlineColor !== undefined) {
    textBox.outlineColor = attrs.outlineColor;
  }
  if (attrs.outlineStyle !== undefined) {
    textBox.outlineStyle = attrs.outlineStyle;
  }
  // Carry anchored-textbox wrap attributes through so the page renderer can
  // build exclusion rects (eigenpal #474).
  if (attrs.displayMode !== undefined) {
    textBox.displayMode = attrs.displayMode;
  }
  if (attrs.cssFloat !== undefined) {
    textBox.cssFloat = attrs.cssFloat;
  }
  if (attrs.wrapType !== undefined) {
    textBox.wrapType = attrs.wrapType;
  }
  if (attrs.wrapText !== undefined) {
    textBox.wrapText = attrs.wrapText;
  }
  if (attrs.distTop !== undefined) {
    textBox.distTop = attrs.distTop;
  }
  if (attrs.distBottom !== undefined) {
    textBox.distBottom = attrs.distBottom;
  }
  if (attrs.distLeft !== undefined) {
    textBox.distLeft = attrs.distLeft;
  }
  if (attrs.distRight !== undefined) {
    textBox.distRight = attrs.distRight;
  }
  if (attrs.position !== undefined) {
    textBox.position = attrs.position;
  }
  if (attrs._docxGroupId !== undefined) {
    setTextBoxGroupId(textBox, attrs._docxGroupId);
  }
  return textBox;
}

function getLastMapKey<K, V>(map: ReadonlyMap<K, V>): K | undefined {
  let lastKey: K | undefined;
  for (const key of map.keys()) {
    lastKey = key;
  }
  return lastKey;
}

/**
 * Convert a ProseMirror document to FlowBlock array.
 *
 * Walks the document tree, converting each node to the appropriate block type.
 * Tracks pmStart/pmEnd positions for each block for click-to-position mapping.
 */
export function toFlowBlocks(doc: PMNode, options: ToFlowBlocksOptions = {}): FlowBlock[] {
  assertValidProseMirrorDocument(doc, "Cannot layout invalid ProseMirror document");

  resetBlockIdCounter();

  const listCounters = options.listCounters ?? new Map<number, number[]>();
  const originalListCounters = options.originalListCounters ?? new Map<number, number[]>();
  const lastAdvancedNumId = getLastMapKey(listCounters);
  const lastAdvancedOriginalNumId = getLastMapKey(originalListCounters);
  const listCounterState: ListCounterState = {
    counters: listCounters,
    abstractCounters: options.listAbstractCounters ?? new Map<number, number[]>(),
    seenLevels: options.listSeenNumIds ?? new Set<string>(),
    ...(lastAdvancedNumId !== undefined ? { lastAdvancedNumId } : {}),
  };
  const originalListCounterState: ListCounterState = {
    counters: originalListCounters,
    abstractCounters: options.originalListAbstractCounters ?? new Map<number, number[]>(),
    seenLevels: options.originalListSeenNumIds ?? new Set<string>(),
    ...(lastAdvancedOriginalNumId !== undefined
      ? { lastAdvancedNumId: lastAdvancedOriginalNumId }
      : {}),
  };

  const opts: FlowConversionOptions = {
    ...options,
    defaultFont: options.defaultFont ?? DEFAULT_FONT,
    defaultSize: options.defaultSize ?? DEFAULT_SIZE,
    listCounters: listCounterState.counters,
    listAbstractCounters: listCounterState.abstractCounters,
    listSeenNumIds: listCounterState.seenLevels,
    originalListCounters: originalListCounterState.counters,
    originalListAbstractCounters: originalListCounterState.abstractCounters,
    originalListSeenNumIds: originalListCounterState.seenLevels,
    listCounterStreams: {
      final: listCounterState,
      original: originalListCounterState,
    },
    numberedRefResults: resolveNumberedRefFields(doc, {
      listCounterState: cloneListCounterState(listCounterState),
      originalListCounterState: cloneListCounterState(originalListCounterState),
    }),
  };

  const blocks: FlowBlock[] = [];
  const offset = 0; // Start at document beginning; kept for clarity.
  void offset;
  let lastSectionMarginsTwips = {
    top: 1440,
    bottom: 1440,
    left: 1440,
    right: 1440,
  };

  let sdtSeq = 0;
  const sdtStack: SdtGroup[] = [];

  /**
   * Stamp the active SDT stack (outer→inner) onto every block produced by the
   * current call. ParagraphBlock/TableBlock accept `sdtGroups`; section breaks
   * and page breaks deliberately do not — they bracket layout, not content.
   */
  const tagBlockWithSdtStack = (block: FlowBlock): void => {
    if (sdtStack.length === 0) {
      return;
    }
    if (block.kind === "paragraph" || block.kind === "table") {
      block.sdtGroups = [...sdtStack];
    }
  };

  const trackedPush = (block: FlowBlock): void => {
    tagBlockWithSdtStack(block);
    blocks.push(block);
  };

  const trailingPageBreakSectionPositions = new Set<number>();
  const consumedPageBreakPositions = new Set<number>();
  const collectTrailingPageBreakSections = (parent: PMNode, contentStart: number): void => {
    let childStart = contentStart;
    for (let index = 0; index < parent.childCount; index += 1) {
      const child = parent.child(index);
      const next = index + 1 < parent.childCount ? parent.child(index + 1) : undefined;
      const paragraphAttrs =
        child.type.name === "paragraph" ? expectParagraphAttrs(child) : undefined;
      if (
        paragraphAttrs &&
        next?.type.name === "pageBreak" &&
        paragraphAttrs._trailingPageBreak === true &&
        (paragraphAttrs._sectionProperties || paragraphAttrs.sectionBreakType)
      ) {
        trailingPageBreakSectionPositions.add(childStart);
        consumedPageBreakPositions.add(childStart + child.nodeSize);
      }
      if (child.type.name === "blockSdt") {
        collectTrailingPageBreakSections(child, childStart + 1);
      }
      childStart += child.nodeSize;
    }
  };
  collectTrailingPageBreakSections(doc, offset);

  // Refactored visit-style traversal so blockSdt can recurse into its
  // children without duplicating the per-block conversion code.
  const visit = (node: PMNode, pos: number): void => {
    switch (node.type.name) {
      case "blockSdt": {
        const attrs = expectBlockSdtAttrs(node);
        sdtSeq += 1;
        const group: SdtGroup = {
          id: `sdt-${sdtSeq}`,
          pmPos: pos,
          sdtType: attrs.sdtType,
        };
        if (attrs.alias) {
          group.alias = attrs.alias;
        }
        if (attrs.tag) {
          group.tag = attrs.tag;
        }
        if (typeof attrs.id === "number") {
          group.sdtId = attrs.id;
        }
        if (attrs.lock) {
          group.lock = attrs.lock;
        }
        if (attrs.showingPlaceholder) {
          group.showingPlaceholder = true;
        }
        if (typeof attrs.checked === "boolean") {
          group.checked = attrs.checked;
        }
        if (attrs.dateFormat) {
          group.dateFormat = attrs.dateFormat;
        }
        if (attrs.listItems) {
          group.listItemsJson = attrs.listItems;
        }

        sdtStack.push(group);
        const startIndex = blocks.length;
        let childOffset = pos + 1; // skip the blockSdt opening token
        for (let i = 0; i < node.childCount; i += 1) {
          const child = node.child(i);
          visit(child, childOffset);
          childOffset += child.nodeSize;
        }
        sdtStack.pop();
        // Stamp first/middle/last/only on the innermost group of each block
        // that was emitted inside this SDT so the painter chrome continues
        // visually across the block sequence. Only paragraph/table blocks
        // carry sdtGroups; section breaks etc. were skipped at tag time.
        const groupBlocks: (ParagraphBlock | TableBlock)[] = [];
        for (let i = startIndex; i < blocks.length; i += 1) {
          const b = blocks[i];
          if (b && (b.kind === "paragraph" || b.kind === "table")) {
            groupBlocks.push(b);
          }
        }
        // We're finalizing the SDT that `group` represents — locate that
        // entry by `pmPos` instead of always taking `at(-1)`. For blocks
        // that sit inside an inner SDT, `at(-1)` is the inner group, and
        // overwriting it would clobber the inner SDT's first/middle/last
        // markers when the outer iterates the same range later. Matching
        // by pmPos keeps the inner positions intact.
        const ourPmPos = group.pmPos;
        for (let i = 0; i < groupBlocks.length; i += 1) {
          const b = groupBlocks[i];
          if (!b || !b.sdtGroups) {
            continue;
          }
          const idx = b.sdtGroups.findIndex((g) => g.pmPos === ourPmPos);
          if (idx === -1) {
            continue;
          }
          let position: NonNullable<SdtGroup["position"]>;
          if (groupBlocks.length === 1) {
            position = "only";
          } else if (i === 0) {
            position = "first";
          } else if (i === groupBlocks.length - 1) {
            position = "last";
          } else {
            position = "middle";
          }
          // Replace just the entry for this SDT, leaving inner/outer
          // sibling entries untouched. (SdtGroup objects are shared
          // across blocks of the same group; copy-on-write here keeps
          // the other blocks' references stable.)
          const next = [...b.sdtGroups];
          const existing = next[idx];
          if (!existing) {
            continue;
          }
          next[idx] = { ...existing, position };
          b.sdtGroups = next;
        }
        return;
      }
      default:
        break;
    }

    switch (node.type.name) {
      case "paragraph": {
        const pmAttrs = expectParagraphAttrs(node);
        const secProps = pmAttrs._sectionProperties as SectionProperties | undefined;
        const hasSectionBreak =
          secProps !== undefined ||
          (pmAttrs.sectionBreakType !== null && pmAttrs.sectionBreakType !== undefined);
        const hasListFormatting =
          (pmAttrs.numPr !== null && pmAttrs.numPr !== undefined) ||
          (pmAttrs.listMarker !== null && pmAttrs.listMarker !== undefined);
        const firstChild = node.firstChild;
        const startsWithColumnBreak =
          firstChild?.type.name === "hardBreak" &&
          expectHardBreakAttrs(firstChild).breakType === "column";
        const isStandaloneColumnBreak = node.childCount === 1 && startsWithColumnBreak;

        if (isStandaloneColumnBreak) {
          const columnBreak: ColumnBreakBlock = {
            kind: "columnBreak",
            id: nextBlockId(),
            pmStart: pos,
            pmEnd: pos + node.nodeSize,
          };
          trackedPush(columnBreak);
        } else if (startsWithColumnBreak && firstChild) {
          const columnBreak: ColumnBreakBlock = {
            kind: "columnBreak",
            id: nextBlockId(),
            pmStart: pos + 1,
            pmEnd: pos + 1 + firstChild.nodeSize,
          };
          trackedPush(columnBreak);

          const paragraph = convertParagraph(node, pos, opts);
          if (paragraph.runs.at(0)?.kind === "lineBreak") {
            paragraph.runs.shift();
          }
          trackedPush(paragraph);
        } else if (node.content.size > 0 || hasListFormatting || !hasSectionBreak) {
          // An empty paragraph carrying w:sectPr is Word's structural section
          // marker; it does not paint an additional blank line. Text-bearing
          // section-ending paragraphs still participate in normal layout.
          trackedPush(convertParagraph(node, pos, opts));
        }

        // Emit section break block if this paragraph ends a section
        if (hasSectionBreak) {
          if (trailingPageBreakSectionPositions.has(pos)) {
            const sourceParagraph = blocks.at(-1);
            if (sourceParagraph?.kind === "paragraph" && sourceParagraph.pmStart === pos) {
              const pageBreak: PageBreakBlock = {
                kind: "pageBreak",
                id: nextBlockId(),
                pmStart: pos + node.nodeSize,
                pmEnd: pos + node.nodeSize + 1,
              };
              trackedPush(pageBreak);
              const sourceAttrs = sourceParagraph.attrs;
              const carrierSpacing = sourceAttrs?.spacing ? { ...sourceAttrs.spacing } : undefined;
              if (carrierSpacing) {
                delete carrierSpacing.before;
              }
              const carrierAttrs: ParagraphAttrs = {
                ...(carrierSpacing ? { spacing: carrierSpacing } : {}),
                ...(sourceAttrs?.automaticSpacing?.after === true
                  ? { automaticSpacing: { after: true } }
                  : {}),
                ...(sourceAttrs?.spacingExplicit?.after === true
                  ? { spacingExplicit: { after: true } }
                  : {}),
                ...(sourceAttrs?.hasDirectParagraphFormatting === true
                  ? { hasDirectParagraphFormatting: true }
                  : {}),
                ...(sourceAttrs?.hasDirectParagraphMarkFormatting === true
                  ? { hasDirectParagraphMarkFormatting: true }
                  : {}),
                ...(sourceAttrs?.snapToGrid !== undefined
                  ? { snapToGrid: sourceAttrs.snapToGrid }
                  : {}),
                ...(sourceAttrs?.documentGridLinePitch !== undefined
                  ? { documentGridLinePitch: sourceAttrs.documentGridLinePitch }
                  : {}),
                ...(sourceAttrs?.defaultFontSize !== undefined
                  ? { defaultFontSize: sourceAttrs.defaultFontSize }
                  : {}),
                ...(sourceAttrs?.defaultFontFamily !== undefined
                  ? { defaultFontFamily: sourceAttrs.defaultFontFamily }
                  : {}),
                ...(sourceAttrs?.defaultAlternateFontFamily !== undefined
                  ? { defaultAlternateFontFamily: sourceAttrs.defaultAlternateFontFamily }
                  : {}),
              };
              const carrier: ParagraphBlock = {
                kind: "paragraph",
                id: nextBlockId(),
                runs: [],
                attrs: carrierAttrs,
                pmStart: pos + node.nodeSize,
                pmEnd: pos + node.nodeSize,
              };
              trackedPush(carrier);
            }
          }
          const sectionBreak: SectionBreakBlock = {
            kind: "sectionBreak",
            id: nextBlockId(),
          };
          const breakType = secProps?.sectionStart ?? pmAttrs.sectionBreakType;
          if (breakType) {
            sectionBreak.type = breakType as NonNullable<SectionBreakBlock["type"]>;
          }

          if (secProps) {
            sectionBreak.pageNumbering = getPageNumbering(secProps);
            const documentGridLinePitchTwips = resolveDocumentGridLinePitch(secProps.docGrid);
            if (documentGridLinePitchTwips !== undefined) {
              sectionBreak.documentGridLinePitchTwips = documentGridLinePitchTwips;
            }
            // Populate page size
            if (secProps.pageWidth || secProps.pageHeight) {
              sectionBreak.pageSize = {
                w: twipsToPixels(secProps.pageWidth ?? 12_240),
                h: twipsToPixels(secProps.pageHeight ?? 15_840),
              };
            }
            // Populate margins
            if (
              secProps.marginTop !== undefined ||
              secProps.marginBottom !== undefined ||
              secProps.marginLeft !== undefined ||
              secProps.marginRight !== undefined
            ) {
              lastSectionMarginsTwips = {
                top: secProps.marginTop ?? lastSectionMarginsTwips.top,
                bottom: secProps.marginBottom ?? lastSectionMarginsTwips.bottom,
                left: secProps.marginLeft ?? lastSectionMarginsTwips.left,
                right: secProps.marginRight ?? lastSectionMarginsTwips.right,
              };
              sectionBreak.margins = {
                top: twipsToPixels(lastSectionMarginsTwips.top),
                bottom: twipsToPixels(lastSectionMarginsTwips.bottom),
                left: twipsToPixels(lastSectionMarginsTwips.left),
                right: twipsToPixels(lastSectionMarginsTwips.right),
              };
              if (secProps.headerDistance !== undefined) {
                sectionBreak.margins.header = twipsToPixels(secProps.headerDistance);
              }
              if (secProps.footerDistance !== undefined) {
                sectionBreak.margins.footer = twipsToPixels(secProps.footerDistance);
              }
            }
            // Populate columns
            const columns = getColumns(secProps);
            if (columns) {
              sectionBreak.columns = columns;
            }
          }

          trackedPush(sectionBreak);
        }
        break;
      }

      case "table":
        trackedPush(convertTable(node, pos, opts));
        break;

      case "image":
        // Standalone image block (if not inline)
        trackedPush(convertImage(node, pos, opts.pageContentHeight));
        break;

      case "textBox":
        trackedPush(convertTextBoxNode(node, pos, opts));
        break;

      case "horizontalRule":
      case "pageBreak": {
        if (consumedPageBreakPositions.has(pos)) {
          break;
        }
        const pb: PageBreakBlock = {
          kind: "pageBreak",
          id: nextBlockId(),
          pmStart: pos,
          pmEnd: pos + node.nodeSize,
        };
        trackedPush(pb);
        break;
      }
      default:
        break;
    }
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  doc.forEach((node, nodeOffset) => {
    visit(node, offset + nodeOffset);
  });

  reserveLeadingEmptyOutlineHeight(blocks);
  suppressFinalEmptyParagraphAfterTable(blocks);
  suppressFinalParagraphInRepeatedEmptySuffix(blocks);
  const mergedBlocks = mergeRunInParagraphs(blocks);
  const griddedBlocks = applySectionDocumentGrid(
    mergedBlocks,
    opts.finalSectionDocumentGridLinePitchTwips,
  );
  return groupParagraphFrames(griddedBlocks, nextBlockId);
}

function applySectionDocumentGrid(
  blocks: FlowBlock[],
  finalLinePitchTwips: number | undefined,
): FlowBlock[] {
  const result = [...blocks];
  let sectionStart = 0;

  const stampSection = (end: number, linePitchTwips: number | undefined): void => {
    if (linePitchTwips === undefined || linePitchTwips <= 0) {
      return;
    }
    const linePitch = twipsToPixels(linePitchTwips);
    for (let index = sectionStart; index < end; index += 1) {
      const block = result[index];
      if (block?.kind !== "paragraph") {
        continue;
      }
      result[index] = {
        ...block,
        attrs: { ...block.attrs, documentGridLinePitch: linePitch },
      };
    }
  };

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (block?.kind !== "sectionBreak") {
      continue;
    }
    stampSection(index, block.documentGridLinePitchTwips);
    sectionStart = index + 1;
  }
  stampSection(blocks.length, finalLinePitchTwips);
  return result;
}

/**
 * Merge consecutive paragraph blocks where the first carries
 * `runInWithNext` (`<w:specVanish/>` on the paragraph mark).
 *
 * Word's run-in heading feature renders the next paragraph inline on
 * the same line, so for layout we collapse the pair into one
 * ParagraphBlock with combined runs. The merged block keeps the first
 * paragraph's attrs (heading formatting, list marker, indent) and
 * extends pmEnd to the second paragraph's range so click-to-position
 * resolution still maps both ranges back to body content.
 *
 * Chains: runInWithNext on the merged block is dropped because the
 * second paragraph's mark wasn't `specVanish`. If a chain of
 * specVanish paragraphs needs collapsing (rare in practice), the loop
 * naturally handles it by re-inspecting the merged block's flag (we
 * preserve runInWithNext only when the second paragraph itself has
 * specVanish).
 */
function sdtGroupStacksEqual(a: SdtGroup[] | undefined, b: SdtGroup[] | undefined): boolean {
  // pmPos is unique per SDT instance within a single toFlowBlocks call, so
  // comparing the pmPos stacks is enough to tell "same membership" vs.
  // "different membership" without copying the rest of the group payload.
  const lenA = a?.length ?? 0;
  const lenB = b?.length ?? 0;
  if (lenA !== lenB) {
    return false;
  }
  if (lenA === 0) {
    return true;
  }
  for (let i = 0; i < lenA; i++) {
    if (a?.[i]?.pmPos !== b?.[i]?.pmPos) {
      return false;
    }
  }
  return true;
}

function mergeRunInParagraphs(blocks: FlowBlock[]): FlowBlock[] {
  const out: FlowBlock[] = [];
  for (let i = 0; i < blocks.length; i++) {
    let current = blocks[i];
    if (!current) {
      continue;
    }
    // Chain merge: keep folding consecutive paragraphs while the
    // *current* (possibly already-merged) block carries
    // `runInWithNext` and the next block is also a paragraph. Per
    // ECMA-376 §17.3.1.32 and Word's behaviour, a sequence of
    // `<w:specVanish/>` paragraphs flows inline through the first
    // body paragraph that lacks it (Codex PR #258 review).
    while (
      current.kind === "paragraph" &&
      (current as ParagraphBlock).attrs?.runInWithNext &&
      i + 1 < blocks.length
    ) {
      const next = blocks[i + 1];
      if (!next || next.kind !== "paragraph") {
        break;
      }
      const a = current as ParagraphBlock;
      const b = next as ParagraphBlock;
      // Stop merging when the two paragraphs sit in different SDT stacks.
      // The merged ParagraphBlock would inherit only `a`'s sdtGroups via
      // spread, so a `<w:specVanish/>` adjacent to a paragraph across an
      // SDT boundary would either claim outside text as part of the SDT
      // or strip SDT membership from inside text — chrome and widget
      // click targets would line up against the wrong content range.
      if (!sdtGroupStacksEqual(a.sdtGroups, b.sdtGroups)) {
        break;
      }
      const mergedAttrs: ParagraphAttrs = { ...a.attrs };
      // Heading typically has no spaceAfter; the body's spaceAfter
      // governs the merged paragraph's trailing gap.
      if (b.attrs?.spacing?.after !== undefined) {
        mergedAttrs.spacing = {
          ...mergedAttrs.spacing,
          after: b.attrs.spacing.after,
        };
      }
      // Carry forward `runInWithNext` only if the *consumed* second
      // paragraph itself was specVanish — the while condition above
      // then triggers another fold against the paragraph after it.
      if (b.attrs?.runInWithNext) {
        mergedAttrs.runInWithNext = true;
      } else {
        delete mergedAttrs.runInWithNext;
      }
      const merged: ParagraphBlock = {
        ...a,
        runs: [...a.runs, ...b.runs],
        attrs: mergedAttrs,
      };
      const mergedPmEnd = b.pmEnd ?? a.pmEnd;
      if (mergedPmEnd !== undefined) {
        merged.pmEnd = mergedPmEnd;
      }
      current = merged;
      i += 1; // consumed `next`; fold further if the merged block still has the flag
    }
    out.push(current);
  }
  return out;
}
