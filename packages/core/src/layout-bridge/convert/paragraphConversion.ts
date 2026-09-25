/**
 * Converts a ProseMirror paragraph node to a ParagraphBlock, including
 * paragraph visibility and payload predicates.
 */

import type { Node as PMNode } from "prosemirror-model";
import { panic } from "better-result";
import { statesNoBorder } from "@stll/docx-core/model";
import { getFontAlternate } from "../../fonts/fontAlternates";
import type { ParagraphBlock, Run, ParagraphAttrs } from "../../layout-engine/types";
import { setParagraphFrame } from "../../layout-engine/paragraphFrame";
import {
  resolveParagraphMarkFormatting,
  resolveParagraphMarkOwnFormatting,
} from "./paragraphMarkFormatting";
import { expectParagraphAttrs } from "../../prosemirror/attrs";
import { expectTextBoxAnchorAttrs } from "../../prosemirror/textBoxAnchorAttrs";
import type { ParagraphAttrs as PMParagraphAttrs } from "../../prosemirror/schema/nodes";
import type { TextFormatting } from "../../types/document";
import { twipsToPixels, nextBlockId } from "./flowConversionShared";
import type { FlowConversionOptions } from "./flowConversionShared";
import { resolveWesternThemeFont } from "./textFormattingConversion";
import { paragraphToRuns } from "./paragraphRuns";
import type { PageBreakRunProjection } from "./paragraphRuns";
import { convertParagraphAttrs } from "./paragraphAttrs";

const DETACHED_WATERMARK_HOST_ATTR = "_detachedWatermarkHost";

const expectDetachedWatermarkHostAttr = (attrs: Readonly<Record<string, unknown>>): boolean => {
  const value = Reflect.get(attrs, DETACHED_WATERMARK_HOST_ATTR);
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value !== "boolean") {
    panic(
      "Invalid ProseMirror detached watermark host attrs:\nparagraph.attrs._detachedWatermarkHost: Expected a boolean.",
    );
  }
  return value;
};

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

function hasOnlyHiddenTextRuns(runs: Run[]): boolean {
  return runs.length > 0 && runs.every((run) => run.kind === "text" && run.hidden === true);
}

export function hasVisibleParagraphPayload(attrs: ParagraphAttrs): boolean {
  return (
    (attrs.listMarker !== undefined && !attrs.listMarkerHidden) ||
    attrs.borders?.top !== undefined ||
    attrs.borders?.bottom !== undefined ||
    attrs.borders?.left !== undefined ||
    attrs.borders?.right !== undefined ||
    attrs.borders?.between !== undefined ||
    attrs.borders?.bar !== undefined ||
    attrs.shading !== undefined
  );
}

/** Whether any `w:pBdr` side draws a rule (`none`/`nil` sides draw nothing). */
export function drawsParagraphBorder(borders: PMParagraphAttrs["borders"]): boolean {
  if (!borders) {
    return false;
  }
  return [
    borders.top,
    borders.bottom,
    borders.left,
    borders.right,
    borders.between,
    borders.bar,
  ].some((border) => border?.style !== undefined && !statesNoBorder(border.style));
}

export function convertParagraph(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
  pageBreaks?: PageBreakRunProjection[],
): ParagraphBlock {
  const pmAttrs = expectParagraphAttrs(node);
  const runs = paragraphToRuns(node, startPos, options, pageBreaks);
  const attrs = convertParagraphAttrs(pmAttrs, {
    theme: options.theme,
    fontAlternates: options.fontAlternates,
    listCounterStreams: options.listCounterStreams,
    defaultTabStopTwips: options.defaultTabStopTwips,
    paragraphMarkFormatting: () => resolveParagraphMarkFormatting(pmAttrs, options.styleResolver),
  });
  if (options.numberingTabIgnoresIndent && attrs.listMarker !== undefined) {
    attrs.listNumberingTabIgnoresIndent = true;
  }
  if (options.defaultFont !== undefined) {
    attrs.defaultFontFamily ??= options.defaultFont;
  }
  if (options.defaultSize !== undefined) {
    attrs.defaultFontSize ??= options.defaultSize;
  }
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
    const directParagraphMarkFormatting = pmAttrs._originalFormatting?.runProperties;
    if (
      directParagraphMarkFormatting &&
      Object.values(directParagraphMarkFormatting).some(
        (value) => value !== undefined && value !== null,
      )
    ) {
      attrs.hasDirectParagraphMarkFormatting = true;
    }
    // The empty line takes the mark's own size and face, including those of a
    // character style the mark names in `w:rStyle`.
    const paragraphMarkFormatting = resolveParagraphMarkOwnFormatting(
      directParagraphMarkFormatting,
      options.styleResolver,
    );
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
  const isFullyHiddenParagraph =
    defaultTextFormatting?.hidden === true && (runs.length === 0 || hasOnlyHiddenTextRuns(runs));
  if (isFullyHiddenParagraph && attrs.listMarker !== undefined) {
    attrs.listMarkerHidden = true;
  }
  if (isFullyHiddenParagraph) {
    attrs.suppressEmptyParagraphHeight = true;
  }
  if (runs.length === 0 && expectDetachedWatermarkHostAttr(node.attrs)) {
    attrs.suppressEmptyParagraphHeight = false;
  }
  if (
    runs.length === 0 &&
    pmAttrs._pageBreakCarrier === true &&
    !hasVisibleParagraphPayload(attrs)
  ) {
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
  node.descendants((child) => {
    if (child.type.name !== "textBoxAnchor") {
      return true;
    }
    options.textBoxAnchorBlockIds.set(expectTextBoxAnchorAttrs(child).anchorId, block.id);
    return false;
  });
  return block;
}
