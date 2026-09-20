/**
 * Selection State Utilities
 *
 * Extracts selection state from ProseMirror for toolbar integration.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import {
  expectFontFamilyMarkAttrs,
  expectFontSizeMarkAttrs,
  expectHighlightMarkAttrs,
  expectParagraphAttrs,
  expectStrikeMarkAttrs,
  expectTextColorMarkAttrs,
  expectUnderlineMarkAttrs,
} from "./attrs";
import { directionIsRtl } from "./paragraphDirection";
import { collectMarksInRange } from "./selectionMarks";
import type { TextFormatting, ParagraphFormatting } from "../types/document";
import { FONT_THEME_VALUES } from "../types/documentEnumValues";

// ============================================================================
// TYPES
// ============================================================================

/**
 * Canonical selection projection shared by toolbar and selection-change APIs.
 * Adapter-specific context belongs on top of this snapshot rather than in a
 * second formatting extractor.
 */
export type SelectionSnapshot = {
  /** Whether there's an active selection (not just cursor) */
  hasSelection: boolean;
  /** Whether selection spans multiple paragraphs */
  isMultiParagraph: boolean;
  /** Current text formatting at selection/cursor */
  textFormatting: TextFormatting;
  /** Current paragraph formatting */
  paragraphFormatting: ParagraphFormatting;
  /** Current paragraph style ID (e.g., 'Heading1', 'Normal') */
  styleId: string | null;
  /** Start paragraph index */
  startParagraphIndex: number;
  /** End paragraph index */
  endParagraphIndex: number;
};

/** Compatibility name for the public toolbar selection contract. */
export type SelectionState = SelectionSnapshot;

type FontTheme = NonNullable<NonNullable<TextFormatting["fontFamily"]>["asciiTheme"]>;

const isFontTheme = (value: string | undefined): value is FontTheme =>
  value !== undefined && FONT_THEME_VALUES.some((theme) => theme === value);

// ============================================================================
// FUNCTIONS
// ============================================================================

/**
 * Extract the canonical selection snapshot from editor state.
 * Used by both the toolbar callback and the selection tracker plugin.
 */
export function extractSelectionSnapshot(state: EditorState): SelectionSnapshot {
  const { selection, doc } = state;
  const { from, to, empty } = selection;
  const $from = doc.resolve(from);

  const { startParagraphIndex, endParagraphIndex } = paragraphIndexes(doc, from, to);
  const paragraph = $from.parent;
  const paragraphAttrs =
    paragraph.type.name === "paragraph" ? expectParagraphAttrs(paragraph) : null;

  const marks = selectionMarks(state);
  let textFormatting = extractTextFormatting(marks);
  const isEmptyParagraph =
    paragraph.type.name === "paragraph" && paragraph.textContent.length === 0;

  if (isEmptyParagraph && marks.length === 0 && paragraphAttrs?.defaultTextFormatting) {
    textFormatting = { ...paragraphAttrs.defaultTextFormatting };
  }

  const paragraphFormatting: ParagraphFormatting = {};
  let styleId: string | null = null;

  if (paragraphAttrs) {
    if (paragraphAttrs.alignment !== undefined) {
      paragraphFormatting.alignment = paragraphAttrs.alignment;
    }
    if (paragraphAttrs.lineSpacing !== undefined) {
      paragraphFormatting.lineSpacing = paragraphAttrs.lineSpacing;
      if (paragraphAttrs.lineSpacingRule !== undefined) {
        paragraphFormatting.lineSpacingRule = paragraphAttrs.lineSpacingRule;
      }
    }
    if (paragraphAttrs.snapToGrid !== undefined) {
      paragraphFormatting.snapToGrid = paragraphAttrs.snapToGrid;
    }
    if (paragraphAttrs.numPr !== undefined) {
      paragraphFormatting.numPr = paragraphAttrs.numPr;
    }
    if (paragraphAttrs.indentLeft !== undefined) {
      paragraphFormatting.indentLeft = paragraphAttrs.indentLeft;
    }
    if (paragraphAttrs.indentRight !== undefined) {
      paragraphFormatting.indentRight = paragraphAttrs.indentRight;
    }
    if (paragraphAttrs.indentFirstLine !== undefined) {
      paragraphFormatting.indentFirstLine = paragraphAttrs.indentFirstLine;
    }
    if (paragraphAttrs.hangingIndent !== undefined) {
      paragraphFormatting.hangingIndent = paragraphAttrs.hangingIndent;
    }
    if (paragraphAttrs.tabs !== undefined) {
      paragraphFormatting.tabs = paragraphAttrs.tabs;
    }
    if (directionIsRtl(paragraphAttrs.direction)) {
      paragraphFormatting.bidi = true;
    }
    if (paragraphAttrs.styleId !== undefined) {
      styleId = paragraphAttrs.styleId;
    }
  }

  return {
    hasSelection: !empty,
    isMultiParagraph: startParagraphIndex !== endParagraphIndex,
    textFormatting,
    paragraphFormatting,
    styleId,
    startParagraphIndex,
    endParagraphIndex,
  };
}

/** Compatibility entry point used by existing React and Vue integrations. */
export function extractSelectionState(state: EditorState): SelectionState | null {
  return extractSelectionSnapshot(state);
}

function paragraphIndexes(
  doc: PMNode,
  from: number,
  to: number,
): {
  startParagraphIndex: number;
  endParagraphIndex: number;
} {
  let startParagraphIndex = 0;
  let endParagraphIndex = 0;

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  doc.forEach((_node, offset, index) => {
    if (offset <= from) {
      startParagraphIndex = index;
    }
    if (offset <= to) {
      endParagraphIndex = index;
    }
  });

  return { startParagraphIndex, endParagraphIndex };
}

function selectionMarks(state: EditorState): readonly Mark[] {
  const { selection, doc } = state;
  const { from, to, empty } = selection;
  return empty
    ? state.storedMarks || selection.$from.marks()
    : collectMarksInRange({ doc, from, to });
}

function extractTextFormatting(marks: readonly Mark[]): TextFormatting {
  const formatting: TextFormatting = {};

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
        if (attrs.style !== "none") {
          formatting.underline = {
            style: attrs.style ?? "single",
            ...(attrs.color !== undefined ? { color: attrs.color } : {}),
          };
        }
        break;
      }
      case "strike": {
        const attrs = expectStrikeMarkAttrs(mark);
        if (attrs.double) {
          formatting.doubleStrike = true;
        } else {
          formatting.strike = true;
        }
        break;
      }
      case "textColor": {
        const attrs = expectTextColorMarkAttrs(mark);
        formatting.color = {
          ...(attrs.rgb !== undefined ? { rgb: attrs.rgb } : {}),
          ...(attrs.themeColor !== undefined ? { themeColor: attrs.themeColor } : {}),
          ...(attrs.themeTint !== undefined ? { themeTint: attrs.themeTint } : {}),
          ...(attrs.themeShade !== undefined ? { themeShade: attrs.themeShade } : {}),
        };
        break;
      }
      case "highlight":
        formatting.highlight = expectHighlightMarkAttrs(mark).color;
        break;
      case "fontSize":
        formatting.fontSize = expectFontSizeMarkAttrs(mark).size;
        break;
      case "fontFamily": {
        const attrs = expectFontFamilyMarkAttrs(mark);
        formatting.fontFamily = {
          ...(attrs.ascii !== undefined ? { ascii: attrs.ascii } : {}),
          ...(attrs.hAnsi !== undefined ? { hAnsi: attrs.hAnsi } : {}),
          ...(attrs.eastAsia !== undefined ? { eastAsia: attrs.eastAsia } : {}),
          ...(attrs.cs !== undefined ? { cs: attrs.cs } : {}),
          ...(attrs.hint !== undefined ? { hint: attrs.hint } : {}),
          ...(isFontTheme(attrs.asciiTheme) ? { asciiTheme: attrs.asciiTheme } : {}),
          ...(attrs.hAnsiTheme !== undefined ? { hAnsiTheme: attrs.hAnsiTheme } : {}),
          ...(attrs.eastAsiaTheme !== undefined ? { eastAsiaTheme: attrs.eastAsiaTheme } : {}),
          ...(attrs.csTheme !== undefined ? { csTheme: attrs.csTheme } : {}),
        };
        break;
      }
      case "superscript":
        formatting.vertAlign = "superscript";
        break;
      case "subscript":
        formatting.vertAlign = "subscript";
        break;
      default:
        break;
    }
  }

  return formatting;
}
