/**
 * Converts ProseMirror paragraph attributes to layout ParagraphAttrs:
 * alignment, spacing, indentation, borders, tab stops, and list markers.
 */

import { paragraphNumberingLevel, paragraphNumberingReferenceId } from "@stll/docx-core/model";
import { bulletMarkerFontName, convertBulletToUnicode } from "../../docx/bulletMarkers";
import { getFontAlternate, type FontAlternates } from "../../fonts/fontAlternates";
import type { ParagraphAttrs, TabAlignment, TabStop } from "../../layout-engine/types";
import { mergeParagraphTabStops } from "../../utils/paragraphFormattingMerge";
import { autospacingMatchesBase } from "../../prosemirror/autospacingBase";
import { directionIsRtl, directionToBidi } from "../../prosemirror/paragraphDirection";
import type { ParagraphAttrs as PMParagraphAttrs } from "../../prosemirror/schema/nodes";
import { advanceVisibleListMarker, type ListCounterStreams } from "../../prosemirror/listMarker";
import type {
  ParagraphAlignment,
  TabStopAlignment,
  Theme,
  TextFormatting,
} from "../../types/document";
import { AUTO_PARAGRAPH_SPACING_PX } from "../../utils/units";
import { twipsToPixels } from "./flowConversionShared";
import { resolveWesternThemeFont, listMarkerFormattingFor } from "./textFormattingConversion";
import {
  applyMarkerAllCaps,
  toListMarkerRevision,
  isAddedNumberingChange,
  isRemovedNumberingChange,
  isChangedNumberingChange,
  applyDeletedListMarkerAttrs,
} from "./listMarkers";
import type { ListPropertyChange, ListPropertyFormatting } from "./listMarkers";
import { convertBorderSpecToLayout } from "./flowBorders";

type ConvertParagraphAttrsOptions = {
  theme: Theme | null | undefined;
  fontAlternates: FontAlternates | undefined;
  listCounterStreams: ListCounterStreams;
  defaultTabStopTwips: number | undefined;
  /**
   * Resolved paragraph-mark run properties, the base of the list marker's
   * typography. Lazy: only paragraphs that paint a marker resolve them.
   */
  paragraphMarkFormatting: () => TextFormatting | undefined;
};

type FlowAlignment = NonNullable<ParagraphAttrs["alignment"]>;

/**
 * Every `ST_Jc` member, and the flow alignment that paints it.
 *
 * `start` and `end` stay unresolved here: they name an edge of the writing
 * direction, not a side of the page, so which of `left` and `right` paints
 * them depends on the paragraph's direction. `numTab` has no flow alignment at
 * all — it aligns the paragraph to the list number's tab stop, which the flow
 * engine does not model — so it falls back to the start edge, which is where a
 * numbered paragraph without one sits.
 */
const FLOW_ALIGNMENT_BY_PARAGRAPH_ALIGNMENT = {
  start: "start",
  end: "end",
  numTab: "start",
  left: "left",
  center: "center",
  right: "right",
  both: "justify",
  distribute: "justify",
  mediumKashida: "justify",
  highKashida: "justify",
  lowKashida: "justify",
  thaiDistribute: "justify",
} as const satisfies Record<ParagraphAlignment, FlowAlignment | "start" | "end">;

/** Resolve a logical alignment against the direction the paragraph runs in. */
const resolveFlowAlignment = (
  alignment: ParagraphAlignment,
  rightToLeft: boolean,
): FlowAlignment => {
  const flow = FLOW_ALIGNMENT_BY_PARAGRAPH_ALIGNMENT[alignment];
  switch (flow) {
    case "start":
      return rightToLeft ? "right" : "left";
    case "end":
      return rightToLeft ? "left" : "right";
    default:
      return flow;
  }
};

export function convertParagraphAttrs(
  pmAttrs: PMParagraphAttrs,
  {
    theme,
    fontAlternates,
    listCounterStreams,
    defaultTabStopTwips,
    paragraphMarkFormatting,
  }: ConvertParagraphAttrsOptions,
): ParagraphAttrs {
  const attrs: ParagraphAttrs = {};

  if (pmAttrs.alignment) {
    attrs.alignment = resolveFlowAlignment(pmAttrs.alignment, directionIsRtl(pmAttrs.direction));
  }

  if (pmAttrs.outlineLevel !== undefined) {
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
    // direct formatting, document defaults, resolved paragraph/table styles,
    // and automatic spacing. Layout consumes the combined provenance;
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
  if (
    paragraphNumberingReferenceId(pmAttrs.numPr) !== undefined &&
    indentLeft === undefined &&
    indentFirstLine === undefined
  ) {
    // Fallback: calculate indentation based on level
    // An authored first-line or hanging position is already a complete list
    // marker anchor. Adding a synthetic left indent would shift that anchor a
    // second time, while tab stops still resolve from the paragraph margin.
    // Each level indents 0.5 inch (720 twips) more
    const level = paragraphNumberingLevel(pmAttrs.numPr) ?? 0;
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

  // Tab stops. A numbering level's `w:pPr/w:tabs` (§17.9.23) sit beneath the
  // paragraph's own and its style's stops.
  const tabs = mergeParagraphTabStops(
    pmAttrs.listLevelTabs ?? undefined,
    pmAttrs.tabs ?? undefined,
  );
  if (tabs && tabs.length > 0) {
    const rightToLeft = directionIsRtl(pmAttrs.direction);
    attrs.tabs = tabs.map((tab) => {
      const tabStop: TabStop = {
        val: resolveTabAlignment(tab.alignment, rightToLeft),
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
    attrs.numPr = pmAttrs.numPr;

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
      ? convertBulletToUnicode(
          pmAttrs.listMarker,
          bulletMarkerFontName(pmAttrs.listMarkerFormatting),
        )
      : pmAttrs.listMarker;
  }
  if (pmAttrs.listIsBullet !== undefined) {
    attrs.listIsBullet = pmAttrs.listIsBullet;
  }
  if (pmAttrs.listMarkerHidden) {
    attrs.listMarkerHidden = true;
  }
  if (attrs.listMarker !== undefined || pmAttrs.listMarkerFormatting) {
    const listMarkerFormatting = listMarkerFormattingFor(
      pmAttrs.listMarkerFormatting,
      paragraphMarkFormatting(),
      theme,
      fontAlternates,
    );
    if (listMarkerFormatting) {
      attrs.listMarkerFormatting = listMarkerFormatting;
    }
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
      applyDeletedListMarkerAttrs(
        attrs,
        numberingRemovedChange,
        undefined,
        theme,
        fontAlternates,
        paragraphMarkFormatting,
      );
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
    // `w:szCs` sizes complex-script characters only, so the paragraph mark
    // takes it when the mark itself is complex script: in a right-to-left
    // paragraph or under `w:rtl` / `w:cs`.
    if (
      attrs.listMarker !== undefined &&
      !attrs.listMarkerHidden &&
      (attrs.bidi === true || dtf.rtl === true || dtf.cs === true) &&
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
 * Every `ST_TabJc` member, and the layout engine's name for it.
 *
 * The engine's `start` and `end` are its names for the left and right edges of
 * the line, not direction-aware ones, so `w:tab w:val="left"` maps to `start`
 * unconditionally. `ST_TabJc`'s own `start` and `end` *are* direction-aware
 * and stay unresolved here, under names the engine has no member for.
 * `num`, the tab a numbered paragraph's text hangs from, has no engine
 * alignment of its own and falls back to the line's start edge.
 */
const TAB_ALIGNMENT_BY_TAB_STOP_ALIGNMENT = {
  start: "logicalStart",
  end: "logicalEnd",
  num: "logicalStart",
  left: "start",
  right: "end",
  center: "center",
  decimal: "decimal",
  bar: "bar",
  clear: "clear",
} as const satisfies Record<TabStopAlignment, TabAlignment | "logicalStart" | "logicalEnd">;

/** Resolve a logical tab alignment against the direction the paragraph runs in. */
const resolveTabAlignment = (alignment: TabStopAlignment, rightToLeft: boolean): TabAlignment => {
  const mapped = TAB_ALIGNMENT_BY_TAB_STOP_ALIGNMENT[alignment];
  switch (mapped) {
    case "logicalStart":
      return rightToLeft ? "end" : "start";
    case "logicalEnd":
      return rightToLeft ? "start" : "end";
    default:
      return mapped;
  }
};
