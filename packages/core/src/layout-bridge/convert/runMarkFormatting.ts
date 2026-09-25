/**
 * Derives layout RunFormatting from the ProseMirror marks on a text node:
 * mark paint dispositions, bidi wrapper direction, and per-mark formatting.
 */

import type { Mark } from "prosemirror-model";
import { panic } from "better-result";
import type { UnderlineStyle } from "@stll/docx-core/model";
import { getFontAlternate, type FontAlternates } from "../../fonts/fontAlternates";
import type { RunFormatting } from "../../layout-engine/types";
import { setHyperlinkInstanceIndex } from "../../layout-engine/measure/hyperlinkInstance";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import {
  expectCharacterSpacingMarkAttrs,
  expectCommentMarkAttrs,
  expectEmphasisMarkAttrs,
  expectFontFamilyMarkAttrs,
  expectLanguageMarkAttrs,
  expectFontSizeMarkAttrs,
  expectFootnoteRefMarkAttrs,
  expectHighlightMarkAttrs,
  expectHyperlinkMarkAttrs,
  expectInlineWrapperMarkAttrs,
  expectRunFormattingOverrideMarkAttrs,
  expectRunShadingMarkAttrs,
  expectTextColorMarkAttrs,
  expectTextEffectMarkAttrs,
  expectTrackedChangeMarkAttrs,
  expectUnderlineMarkAttrs,
} from "../../prosemirror/attrs";
import { runShadingAttrsToShading } from "../../prosemirror/conversion/runShadingMark";
import type { SchemaMarkName } from "../../prosemirror/extensions/markRegistry";
import type {
  InlineWrapperLayer,
  RunFormattingOverrideAttrs,
} from "../../prosemirror/schema/marks";
import type { ColorValue, Theme } from "../../types/document";
import { resolveColor, resolveHighlightToCss } from "../../utils/colorResolver";
import { resolveShadingFill } from "../../utils/formatToStyle";
import { halfPointsToPixels, halfPointsToPoints } from "../../utils/units";
import { twipsToPixels } from "./flowConversionShared";
import { applyRunFormattingOverrides } from "./runFormattingMerge";
import {
  resolveWesternThemeFont,
  resolveComplexScriptThemeFont,
  resolveEastAsiaThemeFont,
  isAutomaticTextColorValue,
} from "./textFormattingConversion";

/**
 * What the painter does with each mark the schema declares.
 *
 * The switch below used to end in `default: break`, so a mark that reached it
 * was silently not painted whether that was the decision or an omission. The
 * table is total over {@link SchemaMarkName}: a mark added to the schema
 * without a paint decision does not compile, and a mark this table says is
 * painted but the switch does not handle panics rather than disappearing.
 */
const MARK_PAINT_DISPOSITIONS = {
  bold: "runFormatting",
  italic: "runFormatting",
  underline: "runFormatting",
  strike: "runFormatting",
  textColor: "runFormatting",
  runShading: "runFormatting",
  highlight: "runFormatting",
  fontSize: "runFormatting",
  fontFamily: "runFormatting",
  language: "runFormatting",
  superscript: "runFormatting",
  subscript: "runFormatting",
  hyperlink: "runFormatting",
  allCaps: "runFormatting",
  smallCaps: "runFormatting",
  footnoteRef: "runFormatting",
  characterSpacing: "runFormatting",
  emboss: "runFormatting",
  imprint: "runFormatting",
  hidden: "runFormatting",
  textShadow: "runFormatting",
  emphasisMark: "runFormatting",
  textOutline: "runFormatting",
  rtl: "runFormatting",
  textEffect: "runFormatting",
  runFormattingOverride: "runFormatting",
  comment: "runFormatting",
  insertion: "runFormatting",
  deletion: "runFormatting",
  inlineWrapper: "runFormatting",
  // Resolved against the style engine by `applyCharacterStyleToggleFormatting`,
  // which reads the same mark list; a branch here would apply it twice.
  characterStyle: "resolvedElsewhere",
  // One authored run's identity and the markup it carried, so a save can
  // rebuild that run. Nothing about it is drawn: an rsid records who edited
  // the text and when, and the `w:rPr` sink holds what no reader took a value
  // from, so nothing in it reached the formatting the painter draws.
  runIdentity: "notPainted",
  // The record of what the run's properties were before the revision. The
  // painter draws the properties the run has now, which the formatting marks
  // beside this one already carry.
  runPropertyChange: "notPainted",
} as const satisfies Record<SchemaMarkName, "runFormatting" | "resolvedElsewhere" | "notPainted">;

const markPaintDisposition = (name: string): string | undefined => {
  const byName: Record<string, string | undefined> = MARK_PAINT_DISPOSITIONS;
  return byName[name];
};

/**
 * Whether a layer states how its content is laid out.
 *
 * A smart tag and a custom-XML wrapper name an element in another vocabulary
 * and say nothing about layout, so the painter draws the text exactly as it
 * would without them. The `switch` is here rather than inline so a kind added
 * to the model has to answer the question.
 */
const paintedDirection = (layer: InlineWrapperLayer): RunFormatting["bidiWrapper"] => {
  switch (layer.kind) {
    case "bidi":
      return layer.direction === undefined
        ? { control: layer.control }
        : { control: layer.control, direction: layer.direction };
    case "smartTag":
    case "customXml":
      return undefined;
    default:
      layer satisfies never;
      return undefined;
  }
};

/**
 * The direction the innermost bidirectional layer of a stack lays its content out in.
 *
 * Innermost wins: a bidirectional wrapper inside another is the one the text is
 * laid out in, and the outer layers are what the save leg rebuilds the nesting
 * from. The layers that state no direction are skipped rather than ending the
 * search, so a smart tag inside a `w:dir` still reads right to left.
 */
const innermostBidiWrapper = (
  stack: readonly InlineWrapperLayer[],
): RunFormatting["bidiWrapper"] => {
  for (const layer of stack.toReversed()) {
    const direction = paintedDirection(layer);
    if (direction !== undefined) {
      return direction;
    }
  }
  return undefined;
};

/**
 * Extract run formatting from ProseMirror marks.
 */
export function extractRunFormatting(
  marks: readonly Mark[],
  theme?: Theme | null,
  fontAlternates?: FontAlternates,
): RunFormatting {
  const formatting: RunFormatting = {};
  let hasNoteRef = false;
  let runFormattingOverride: RunFormattingOverrideAttrs | undefined;

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
          const underlineObj: { style?: UnderlineStyle; color?: string } = {};
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
        if (attrs.hint) {
          formatting.eastAsiaHint = attrs.hint === "eastAsia";
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
        runFormattingOverride = expectRunFormattingOverrideMarkAttrs(mark);
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
        if (attrs._historicalFormatting) {
          formatting.usesHistoricalFormatting = true;
        }
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

      case "inlineWrapper": {
        const bidiWrapper = innermostBidiWrapper(expectInlineWrapperMarkAttrs(mark).stack);
        if (bidiWrapper) {
          formatting.bidiWrapper = bidiWrapper;
        }
        break;
      }

      default:
        if (markPaintDisposition(mark.type.name) === "runFormatting") {
          panic(`Mark ${mark.type.name} is painted as run formatting but has no branch here`);
        }
        break;
    }
  }

  if (runFormattingOverride) {
    // ProseMirror orders marks by schema rank, not semantic ownership. Apply
    // structural cancellations last so an inherited visual mark cannot win.
    applyRunFormattingOverrides(formatting, runFormattingOverride);
  }

  if (hasNoteRef && formatting.subscript) {
    delete formatting.superscript;
  }

  return formatting;
}
