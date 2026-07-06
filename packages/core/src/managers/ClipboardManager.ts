/**
 * ClipboardManager
 *
 * Framework-agnostic DOM-run-extraction surface for clipboard copy/cut.
 * Ported from the upstream docx-editor `managers/ClipboardManager` module.
 *
 * Handles:
 * - DOM selection traversal and run extraction from rendered `.docx-run` nodes
 * - Formatting extraction from computed styles
 *
 * Upstream exposes only these helpers plus the `ClipboardSelection` type (no
 * manager class); this port keeps that surface.
 */

import type { Run, TextFormatting } from "../types/document";

// ============================================================================
// TYPES
// ============================================================================

/** Selection data for clipboard operations. */
export type ClipboardSelection = {
  text: string;
  runs: Run[];
  startParagraphIndex: number;
  startRunIndex: number;
  startOffset: number;
  endParagraphIndex: number;
  endRunIndex: number;
  endOffset: number;
  isMultiParagraph: boolean;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Convert a CSS color string (rgb/rgba/hex) to a 6-char uppercase hex string.
 *
 * NOTE: This parses CSS color *strings*, unlike a numeric-component
 * `rgbToHex(r, g, b)` helper.
 */
export function cssColorToHex(color: string): string | null {
  if (!color || color === "transparent" || color === "inherit") {
    return null;
  }

  if (color.startsWith("#")) {
    return color.slice(1).toUpperCase();
  }

  const rgbMatch = /rgba?\((?<r>\d+),\s*(?<g>\d+),\s*(?<b>\d+)/u.exec(color);
  if (rgbMatch?.groups) {
    const { r, g, b } = rgbMatch.groups;
    if (r === undefined || g === undefined || b === undefined) {
      return null;
    }
    const rHex = Number.parseInt(r, 10).toString(16).padStart(2, "0");
    const gHex = Number.parseInt(g, 10).toString(16).padStart(2, "0");
    const bHex = Number.parseInt(b, 10).toString(16).padStart(2, "0");
    return (rHex + gHex + bHex).toUpperCase();
  }

  return null;
}

/** Extract formatting from an HTML element's computed styles. */
export function extractFormattingFromElement(element: HTMLElement): TextFormatting | undefined {
  const style = window.getComputedStyle(element);
  const formatting: TextFormatting = {};

  // Bold
  if (style.fontWeight === "bold" || Number.parseInt(style.fontWeight, 10) >= 700) {
    formatting.bold = true;
  }

  // Italic
  if (style.fontStyle === "italic") {
    formatting.italic = true;
  }

  // Underline / strikethrough
  // `textDecoration` can resolve to `undefined` when both the shorthand and the
  // `textDecorationLine` longhand are empty/unset (JSDOM and some browsers).
  // Fall back to "" so `.includes()` never runs on `undefined`.
  const textDecoration = style.textDecoration || style.textDecorationLine || "";
  if (textDecoration.includes("underline")) {
    formatting.underline = { style: "single" };
  }
  if (textDecoration.includes("line-through")) {
    formatting.strike = true;
  }

  // Font size (convert px to half-points; 1pt = 1.333px at 96dpi)
  const fontSize = Number.parseFloat(style.fontSize);
  if (!Number.isNaN(fontSize) && fontSize > 0) {
    formatting.fontSize = Math.round((fontSize / 1.333) * 2);
  }

  // Font family
  const fontFamily = style.fontFamily.replace(/["']/gu, "").split(",")[0]?.trim();
  if (fontFamily) {
    formatting.fontFamily = { ascii: fontFamily };
  }

  // Color
  const color = style.color;
  if (color && color !== "rgb(0, 0, 0)") {
    const hex = cssColorToHex(color);
    if (hex) {
      formatting.color = { rgb: hex };
    }
  }

  // Background color
  const bgColor = style.backgroundColor;
  if (bgColor && bgColor !== "transparent" && bgColor !== "rgba(0, 0, 0, 0)") {
    const hex = cssColorToHex(bgColor);
    if (hex) {
      formatting.shading = { fill: { rgb: hex } };
    }
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

/** Get selected text from a run element, considering partial selection. */
function getSelectedTextFromRun(runEl: Node, range: Range): string {
  const runRange = document.createRange();
  runRange.selectNodeContents(runEl);

  const startInRun =
    range.compareBoundaryPoints(Range.START_TO_START, runRange) >= 0 &&
    range.compareBoundaryPoints(Range.START_TO_END, runRange) <= 0;
  const endInRun =
    range.compareBoundaryPoints(Range.END_TO_START, runRange) >= 0 &&
    range.compareBoundaryPoints(Range.END_TO_END, runRange) <= 0;

  if (startInRun && endInRun) {
    return range.toString();
  }
  if (startInRun) {
    const tempRange = document.createRange();
    tempRange.setStart(range.startContainer, range.startOffset);
    tempRange.setEnd(runRange.endContainer, runRange.endOffset);
    return tempRange.toString();
  }
  if (endInRun) {
    const tempRange = document.createRange();
    tempRange.setStart(runRange.startContainer, runRange.startOffset);
    tempRange.setEnd(range.endContainer, range.endOffset);
    return tempRange.toString();
  }
  if (range.intersectsNode(runEl)) {
    return runEl.textContent ?? "";
  }

  return "";
}

/** Find the paragraph element containing a node. */
function findParagraphElement(node: Node): HTMLElement | null {
  let current: Node | null = node;
  while (current) {
    if (current instanceof HTMLElement && current.hasAttribute("data-paragraph-index")) {
      return current;
    }
    current = current.parentNode;
  }
  return null;
}

/** Get selected runs from the current DOM selection. */
export function getSelectionRuns(): Run[] {
  const selection = window.getSelection();
  if (!selection || selection.isCollapsed) {
    return [];
  }

  const runs: Run[] = [];
  const range = selection.getRangeAt(0);

  const container = range.commonAncestorContainer;
  const containerElement =
    container instanceof HTMLElement ? container : container.parentElement;

  if (!containerElement) {
    return runs;
  }

  const runElements = containerElement.querySelectorAll(".docx-run");

  for (const runEl of runElements) {
    if (!(runEl instanceof HTMLElement) || !range.intersectsNode(runEl)) {
      continue;
    }
    const text = getSelectedTextFromRun(runEl, range);
    if (!text) {
      continue;
    }
    const formatting = extractFormattingFromElement(runEl);
    const run: Run = formatting
      ? { type: "run", formatting, content: [{ type: "text", text }] }
      : { type: "run", content: [{ type: "text", text }] };
    runs.push(run);
  }

  if (runs.length === 0) {
    const selectedText = selection.toString();
    if (selectedText) {
      runs.push({ type: "run", content: [{ type: "text", text: selectedText }] });
    }
  }

  return runs;
}

/** Create a ClipboardSelection from the current DOM selection. */
export function createSelectionFromDOM(): ClipboardSelection | null {
  const selection = window.getSelection();
  if (!selection || selection.isCollapsed) {
    return null;
  }

  const runs = getSelectionRuns();
  if (runs.length === 0) {
    return null;
  }

  const text = selection.toString();
  const range = selection.getRangeAt(0);
  const startPara = findParagraphElement(range.startContainer);
  const endPara = findParagraphElement(range.endContainer);

  const startParagraphIndex = startPara
    ? Number.parseInt(startPara.getAttribute("data-paragraph-index") ?? "0", 10)
    : 0;
  const endParagraphIndex = endPara
    ? Number.parseInt(endPara.getAttribute("data-paragraph-index") ?? "0", 10)
    : 0;

  return {
    text,
    runs,
    startParagraphIndex,
    startRunIndex: 0,
    startOffset: range.startOffset,
    endParagraphIndex,
    endRunIndex: 0,
    endOffset: range.endOffset,
    isMultiParagraph: startParagraphIndex !== endParagraphIndex,
  };
}

// Backwards-compatible alias for the CSS-string color converter.
export const rgbToHex = cssColorToHex;
