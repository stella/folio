/**
 * Layout Engine - Main Entry Point
 *
 * Converts blocks + measures into positioned fragments on pages.
 */

import { panic } from "better-result";

import { emuToPixels } from "../utils/units";
import {
  computeKeepNextChains,
  calculateChainHeight,
  getMidChainIndices,
  hasPageBreakBefore,
} from "./keep-together";
import { measuredLineAdvance } from "./lineFlow";
import { FOOTNOTE_SEPARATOR_HEIGHT, createPaginator } from "./paginator";
import { getParagraphFragmentPmRange } from "./paragraphFragmentRange";
import {
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
  isEmptyParagraph,
} from "./paragraphSpacing";
import { buildTableRowBreakInfo, snapRowBreak } from "./tableRowBreak";
import { bandFragmentX, bandTopContentY, isPageFrameRelativeAnchor } from "./textBoxFlow";
import { floatingTextBoxReservesBand } from "./types";
import type {
  FlowBlock,
  Measure,
  Layout,
  LayoutOptions,
  Page,
  PageMargins,
  ColumnLayout,
  ParagraphBlock,
  ParagraphMeasure,
  ParagraphFragment,
  TableBlock,
  TableMeasure,
  TableFragment,
  ImageBlock,
  ImageMeasure,
  ImageFragment,
  TextBoxBlock,
  TextBoxMeasure,
  TextBoxFragment,
  SectionBreakBlock,
} from "./types";

export type SectionLayoutConfig = {
  pageSize: { w: number; h: number };
  margins: PageMargins;
  columns?: ColumnLayout;
};

const DEFAULT_COLUMNS: ColumnLayout = { count: 1, gap: 0 };

export function collectSectionConfigs(
  blocks: FlowBlock[],
  initialConfig: SectionLayoutConfig,
  finalConfig: SectionLayoutConfig,
): {
  configs: SectionLayoutConfig[];
  breakIndices: number[];
} {
  const configs: SectionLayoutConfig[] = [];
  const breakIndices: number[] = [];
  let previousConfig = initialConfig;

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (block?.kind !== "sectionBreak") {
      continue;
    }

    const sectionBreak = block as SectionBreakBlock;
    const config: SectionLayoutConfig = {
      pageSize: sectionBreak.pageSize ?? previousConfig.pageSize,
      margins: sectionBreak.margins ?? previousConfig.margins,
    };
    if (sectionBreak.columns !== undefined) {
      config.columns = sectionBreak.columns;
    }
    configs.push(config);
    breakIndices.push(index);
    previousConfig = configs.at(-1) ?? previousConfig;
  }

  configs.push(finalConfig);
  return { configs, breakIndices };
}

function isPaginationEmptyParagraph(block: ParagraphBlock): boolean {
  return block.runs.every((run) => {
    if (run.kind === "text") {
      return run.text.trim().length === 0;
    }
    return run.kind === "tab" || run.kind === "lineBreak";
  });
}

function pageHasVisibleBodyContent(
  page: Page,
  blocksById: ReadonlyMap<string, FlowBlock>,
): boolean {
  for (const fragment of page.fragments) {
    if (fragment.kind !== "paragraph") {
      return true;
    }
    const block = blocksById.get(String(fragment.blockId));
    if (block?.kind !== "paragraph" || !isPaginationEmptyParagraph(block)) {
      return true;
    }
  }
  return false;
}

function continuesNumberedSequence(previous: FlowBlock | undefined, current: FlowBlock): boolean {
  if (previous?.kind !== "paragraph" || current.kind !== "paragraph") {
    return false;
  }
  const previousNumPr = previous.attrs?.numPr;
  const currentNumPr = current.attrs?.numPr;
  if (previousNumPr?.numId === undefined || currentNumPr?.numId === undefined) {
    return false;
  }
  return previousNumPr.numId === currentNumPr.numId && previousNumPr.ilvl === currentNumPr.ilvl;
}

type NaturalPageAdvance = "none" | "ordinary" | "reflowBoundary";

function pageStartsWithPreviousParagraphContinuation(
  page: Page,
  previousBlock: FlowBlock | undefined,
): boolean {
  if (previousBlock?.kind !== "paragraph") {
    return false;
  }

  return page.fragments.some(
    (fragment) =>
      fragment.kind === "paragraph" &&
      String(fragment.blockId) === String(previousBlock.id) &&
      fragment.continuesFromPrev === true,
  );
}
/**
 * Estimate the height of a short paragraph-only multi-column section so its
 * final page can use Word-style balanced columns. Sections containing tables,
 * floating blocks, or authored breaks keep the normal bottom-up pagination;
 * those need structure-aware balancing rather than a height estimate.
 */
type BalancedParagraphSectionHeightOptions = {
  blocks: FlowBlock[];
  measures: Measure[];
  startIndex: number;
  endIndex: number;
  incomingSpacing: number;
  columnCount: number;
  availableHeight: number;
};

function balancedParagraphSectionHeight({
  blocks,
  measures,
  startIndex,
  endIndex,
  incomingSpacing,
  columnCount,
  availableHeight,
}: BalancedParagraphSectionHeightOptions): number | undefined {
  if (columnCount <= 1 || startIndex >= endIndex || availableHeight <= 0) {
    return undefined;
  }

  const lineUnits: number[] = [];
  let totalHeight = 0;
  let tallestUnit = 0;
  let trailingSpacing = incomingSpacing;
  for (let index = startIndex; index < endIndex; index++) {
    const block = blocks[index];
    const measure = measures[index];
    if (block?.kind !== "paragraph" || measure?.kind !== "paragraph") {
      return undefined;
    }
    if (
      block.attrs?.keepNext === true ||
      block.attrs?.keepLines === true ||
      block.runs.some((run) => run.kind === "text" && run.footnoteRefId !== undefined)
    ) {
      return undefined;
    }

    const leadingSpacing = Math.max(getParagraphSpacingBefore(block), trailingSpacing);
    if (measure.lines.length === 0) {
      if (leadingSpacing > 0) {
        lineUnits.push(leadingSpacing);
        totalHeight += leadingSpacing;
        tallestUnit = Math.max(tallestUnit, leadingSpacing);
      }
    }
    for (let lineIndex = 0; lineIndex < measure.lines.length; lineIndex++) {
      const line = measure.lines[lineIndex];
      if (!line) {
        continue;
      }
      const lineHeight = measuredLineAdvance(line);
      const unitHeight = lineHeight + (lineIndex === 0 ? leadingSpacing : 0);
      lineUnits.push(unitHeight);
      totalHeight += unitHeight;
      tallestUnit = Math.max(tallestUnit, unitHeight);
    }
    if (tallestUnit > availableHeight || totalHeight > availableHeight * columnCount) {
      return undefined;
    }
    trailingSpacing = getParagraphSpacingAfter(block);
  }

  if (totalHeight <= 0) {
    return undefined;
  }

  const columnsNeeded = (targetHeight: number): number => {
    let usedHeight = 0;
    let usedColumns = 1;
    for (const unitHeight of lineUnits) {
      if (usedHeight > 0 && usedHeight + unitHeight > targetHeight) {
        usedColumns += 1;
        usedHeight = unitHeight;
      } else {
        usedHeight += unitHeight;
      }
    }
    return usedColumns;
  };

  let lower = Math.max(tallestUnit, totalHeight / columnCount);
  let upper = Math.min(totalHeight, availableHeight);
  if (columnsNeeded(upper) > columnCount) {
    return undefined;
  }
  for (let iteration = 0; iteration < 32; iteration++) {
    const middle = (lower + upper) / 2;
    if (columnsNeeded(middle) <= columnCount) {
      upper = middle;
    } else {
      lower = middle;
    }
  }
  return Math.ceil(upper * 1000) / 1000;
}

function hasWidowControl(block: ParagraphBlock): boolean {
  return block.attrs?.widowControl !== false;
}

/**
 * Apply contextual spacing suppression (OOXML §17.3.1.9).
 *
 * Contextual spacing applies independently to each paragraph: suppress the
 * current paragraph's spaceAfter when it opts in, and the next paragraph's
 * spaceBefore when it opts in, provided both paragraphs share a style. Two
 * absent style ids both refer to the document's default paragraph style.
 *
 * This mutates the block attrs in-place before layout runs.
 */
export function applyContextualSpacing(blocks: FlowBlock[]): void {
  for (let i = 0; i < blocks.length - 1; i++) {
    const curr = blocks[i]!; // SAFETY: i < blocks.length - 1
    const next = blocks[i + 1]!; // SAFETY: i + 1 < blocks.length

    if (curr.kind !== "paragraph" || next.kind !== "paragraph") {
      continue;
    }

    const currAttrs = curr.attrs;
    const nextAttrs = next.attrs;

    if (currAttrs?.styleId !== nextAttrs?.styleId) {
      continue;
    }
    if (currAttrs?.contextualSpacing && currAttrs.spacing) {
      currAttrs.spacing = { ...currAttrs.spacing, after: 0 };
    }
    if (nextAttrs?.contextualSpacing && nextAttrs.spacing) {
      nextAttrs.spacing = { ...nextAttrs.spacing, before: 0 };
    }
  }

  // Recurse into nested block containers (table cells and text boxes) so
  // contextual spacing is suppressed there too — measure, pagination, and the
  // painter all read the (mutated) paragraph spacing, so they stay consistent.
  // eigenpal/docx-editor#699.
  for (const block of blocks) {
    if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          applyContextualSpacing(cell.blocks);
        }
      }
    } else if (block.kind === "textBox") {
      applyContextualSpacing(block.content);
    }
  }
}

/** Suppress automatic inter-item spacing within one numbered sequence. */
const applyAutomaticListSpacing = (blocks: FlowBlock[]): void => {
  for (let index = 1; index < blocks.length; index++) {
    const previous = blocks[index - 1];
    const current = blocks[index];
    if (
      previous?.kind !== "paragraph" ||
      current?.kind !== "paragraph" ||
      !continuesNumberedSequence(previous, current)
    ) {
      continue;
    }

    if (previous.attrs?.automaticSpacing?.after && previous.attrs.spacing) {
      previous.attrs.spacing = { ...previous.attrs.spacing, after: 0 };
    }
    if (current.attrs?.automaticSpacing?.before && current.attrs.spacing) {
      current.attrs.spacing = { ...current.attrs.spacing, before: 0 };
    }
  }

  for (const block of blocks) {
    if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          applyAutomaticListSpacing(cell.blocks);
        }
      }
    } else if (block.kind === "textBox") {
      applyAutomaticListSpacing(block.content);
    }
  }
};

/**
 * Layout a document: convert blocks + measures into pages with positioned fragments.
 *
 * Algorithm:
 * 1. Walk blocks in order with their corresponding measures
 * 2. For each block, create appropriate fragment(s)
 * 3. Use paginator to manage page/column state
 * 4. Handle page breaks, section breaks, and keepNext chains
 */
export function layoutDocument(
  blocks: FlowBlock[],
  measures: Measure[],
  options: LayoutOptions,
): Layout {
  // Validate input
  if (blocks.length !== measures.length) {
    panic(
      `layoutDocument: expected one measure per block (blocks=${blocks.length}, measures=${measures.length})`,
    );
  }

  // Set up options with defaults
  const pageSize = options.pageSize;
  const baseMargins: PageMargins = {
    top: options.margins.top,
    right: options.margins.right,
    bottom: options.margins.bottom,
    left: options.margins.left,
  };
  if (options.margins.header !== undefined) {
    baseMargins.header = options.margins.header;
  }
  if (options.margins.footer !== undefined) {
    baseMargins.footer = options.margins.footer;
  }

  // Use document margins directly for WYSIWYG fidelity
  // Word uses fixed margins from the document - body content always starts at marginTop
  // If header content extends below marginTop, it overlaps (this matches Word behavior)

  const margins = { ...baseMargins };
  const finalPageSize = options.finalPageSize ?? pageSize;
  const finalMargins = options.finalMargins ?? margins;

  // Calculate content width
  const contentWidth = pageSize.w - margins.left - margins.right;
  if (contentWidth <= 0) {
    panic("layoutDocument: page size and margins yield no content area");
  }

  const bodyConfig: SectionLayoutConfig = {
    pageSize,
    margins,
  };
  if (options.columns !== undefined) {
    bodyConfig.columns = options.columns;
  }
  const finalConfig: SectionLayoutConfig = {
    pageSize: finalPageSize,
    margins: finalMargins,
  };
  const finalColumns = options.finalColumns ?? options.columns;
  if (finalColumns !== undefined) {
    finalConfig.columns = finalColumns;
  }
  const { configs: sectionConfigs, breakIndices } = collectSectionConfigs(
    blocks,
    bodyConfig,
    finalConfig,
  );
  const sectionBreakTypes = [
    ...breakIndices.map((index) => (blocks[index] as SectionBreakBlock).type),
    options.bodyBreakType,
  ];
  const initialConfig = sectionConfigs.at(0) ?? bodyConfig;

  // Create paginator with first section's columns
  const paginator = createPaginator({
    pageSize: initialConfig.pageSize,
    margins: initialConfig.margins,
    ...(options.mirrorMargins === true ? { mirrorMargins: true } : {}),
    ...(options.firstPageMargins !== undefined
      ? { firstPageMargins: options.firstPageMargins }
      : {}),
    ...(options.sectionEvenPageMargins !== undefined
      ? { sectionEvenPageMargins: options.sectionEvenPageMargins }
      : {}),
    columns: initialConfig.columns ?? DEFAULT_COLUMNS,
    ...(options.footnoteReservedHeights !== undefined
      ? { footnoteReservedHeights: options.footnoteReservedHeights }
      : {}),
    ...(options.sectionHeaderFooterRefs !== undefined
      ? { sectionHeaderFooterRefs: options.sectionHeaderFooterRefs }
      : {}),
  });

  // Apply contextual spacing: suppress spaceBefore/spaceAfter between
  // consecutive paragraphs that both have contextualSpacing=true and share
  // the same styleId (OOXML spec 17.3.1.9 / ECMA-376 §17.3.1.9).
  applyContextualSpacing(blocks);
  applyAutomaticListSpacing(blocks);

  // Pre-compute keepNext chains for pagination decisions
  const keepNextChains = computeKeepNextChains(blocks);
  const midChainIndices = getMidChainIndices(keepNextChains);
  const blocksById = new Map<string, FlowBlock>();
  for (const block of blocks) {
    blocksById.set(String(block.id), block);
  }

  // Process each block, tracking section break index with a counter (O(1) per break)
  let sectionIdx = 0;
  // Section page geometry for resolving page/margin-pinned topAndBottom bands.
  // The measure pass (extractFloatingZones) uses the section config, not the
  // page's possibly-different first-page margins, so layout must too or the
  // reserved band and painted box desync. eigenpal #694.
  let activeSectionMarginTop = initialConfig.margins.top;
  let activeSectionPageHeight = initialConfig.pageSize.h;
  let activeSectionMarginBottom = initialConfig.margins.bottom;
  let naturalPageAdvanceSinceRenderedBreak: NaturalPageAdvance = "none";
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]!; // SAFETY: i < blocks.length
    const measure = measures[i]!; // SAFETY: measures.length === blocks.length (validated above)

    // A cached pagination marker records a break at this paragraph, not an
    // ordinal page target: the source does not emit markers for every physical
    // page. Honor the break whenever visible content precedes it on Folio's
    // current page. A structural/natural break that opened an empty page, or a
    // page containing only overflowed empty paragraphs, already satisfies it.
    const hasRenderedPageBreak =
      block.kind === "paragraph" && block.attrs?.renderedPageBreakBefore === true;
    if (hasPageBreakBefore(block)) {
      paginator.forcePageBreak();
    } else if (hasRenderedPageBreak) {
      const state = paginator.getCurrentState();
      const previousParagraphAlreadyCrossedMarker = pageStartsWithPreviousParagraphContinuation(
        state.page,
        blocks[i - 1],
      );
      const naturalAdvanceAlreadyCrossedMarker =
        previousParagraphAlreadyCrossedMarker ||
        naturalPageAdvanceSinceRenderedBreak === "reflowBoundary" ||
        (block.kind === "paragraph" &&
          isPaginationEmptyParagraph(block) &&
          naturalPageAdvanceSinceRenderedBreak !== "none") ||
        (naturalPageAdvanceSinceRenderedBreak !== "none" &&
          continuesNumberedSequence(blocks[i - 1], block));
      if (
        !naturalAdvanceAlreadyCrossedMarker &&
        pageHasVisibleBodyContent(state.page, blocksById)
      ) {
        paginator.forcePageBreak();
      }
    }
    if (hasRenderedPageBreak || hasPageBreakBefore(block)) {
      naturalPageAdvanceSinceRenderedBreak = "none";
    }

    // Handle keepNext chains - if this is a chain start, check if chain fits
    const chain = keepNextChains.get(i);
    if (chain && !midChainIndices.has(i)) {
      const chainHeight = calculateChainHeight(chain, blocks, measures);
      const pageBeforeChainLayout = paginator.getCurrentState().page.number;
      paginator.ensureFits(chainHeight);
      if (paginator.getCurrentState().page.number > pageBeforeChainLayout) {
        // The chain moved as one unit across the cached boundary. A rendered
        // marker immediately after it describes the page Folio just opened.
        naturalPageAdvanceSinceRenderedBreak = "reflowBoundary";
      }
    }

    const pageBeforeBlockLayout = paginator.getCurrentState().page.number;
    switch (block.kind) {
      case "paragraph":
        layoutParagraph(
          block,
          measure as ParagraphMeasure,
          paginator,
          paginator.columnWidth,
          options.footnoteHeightById,
        );
        break;

      case "table":
        if (block.floating) {
          layoutFloatingTable(
            block,
            measure as TableMeasure,
            paginator,
            paginator.getContentWidth(),
          );
        } else {
          layoutTable(block, measure as TableMeasure, paginator, options.footnoteHeightById);
        }
        break;

      case "image":
        layoutImage(block, measure as ImageMeasure, paginator);
        break;

      case "textBox":
        layoutTextBox(block as TextBoxBlock, measure as TextBoxMeasure, {
          paginator,
          sectionMarginTop: activeSectionMarginTop,
          sectionPageHeight: activeSectionPageHeight,
          sectionMarginBottom: activeSectionMarginBottom,
        });
        break;

      case "pageBreak":
        paginator.forcePageBreak();
        break;

      case "columnBreak":
        paginator.forceColumnBreak();
        break;

      case "sectionBreak": {
        // Use the NEXT section's columns; for break type, prefer next section's
        // type but fall back to current break's type (preserves explicit 'continuous')
        const nextSectionConfig = sectionConfigs[sectionIdx + 1] ?? initialConfig;
        const nextType = sectionBreakTypes[sectionIdx + 1] ?? sectionBreakTypes[sectionIdx];
        handleSectionBreak(
          block as SectionBreakBlock,
          paginator,
          nextSectionConfig,
          nextType,
          sectionIdx + 1,
        );
        const nextColumns = nextSectionConfig.columns;
        const nextBreakIndex = breakIndices[sectionIdx + 1] ?? blocks.length;
        const nextBreak = blocks[nextBreakIndex];
        const sectionEndsContinuously =
          nextBreakIndex === blocks.length ||
          (nextBreak?.kind === "sectionBreak" && nextBreak.type === "continuous");
        if (nextColumns && sectionEndsContinuously) {
          const state = paginator.getCurrentState();
          const balancedHeight = balancedParagraphSectionHeight({
            blocks,
            measures,
            startIndex: i + 1,
            endIndex: nextBreakIndex,
            incomingSpacing: state.trailingSpacing,
            columnCount: nextColumns.count,
            availableHeight: paginator.getAvailableHeight(),
          });
          if (balancedHeight !== undefined) {
            state.contentBottom = Math.min(state.contentBottom, state.cursorY + balancedHeight);
          }
        }
        activeSectionMarginTop = nextSectionConfig.margins.top;
        activeSectionPageHeight = nextSectionConfig.pageSize.h;
        activeSectionMarginBottom = nextSectionConfig.margins.bottom;
        sectionIdx++;
        break;
      }
      default:
        break;
    }

    const isVisibleBodyBlock =
      block.kind === "paragraph"
        ? !isEmptyParagraph(block)
        : block.kind !== "pageBreak" &&
          block.kind !== "columnBreak" &&
          block.kind !== "sectionBreak";
    const isStructuralBreak =
      block.kind === "pageBreak" || block.kind === "columnBreak" || block.kind === "sectionBreak";
    if (isStructuralBreak) {
      naturalPageAdvanceSinceRenderedBreak = "none";
    } else if (
      isVisibleBodyBlock &&
      paginator.getCurrentState().page.number > pageBeforeBlockLayout
    ) {
      const previousPage = paginator.pages[pageBeforeBlockLayout - 1];
      const blockContinuedAcrossBoundary = previousPage?.fragments.some(
        ({ blockId }) => blockId === block.id,
      );
      naturalPageAdvanceSinceRenderedBreak = blockContinuedAcrossBoundary
        ? "reflowBoundary"
        : "ordinary";
    }
  }

  // Ensure at least one page exists
  if (paginator.pages.length === 0) {
    paginator.getCurrentState();
  }

  return {
    pageSize,
    pages: paginator.pages,
    ...(options.columns !== undefined ? { columns: options.columns } : {}),
    ...(options.pageGap !== undefined ? { pageGap: options.pageGap } : {}),
  };
}

/**
 * Footnote refs whose run sits in line `[fromRun..toRun]`, with their
 * pre-measured content heights. Empty when the engine isn't tracking
 * dynamic fn demand. Used by `layoutParagraph` to (a) reserve fn
 * height per line and (b) record the IDs on the host page so the
 * painter renders the fn on the page where the ref-bearing line
 * actually landed — even when the paragraph splits across pages
 * (fragment pmStart/pmEnd is paragraph-wide and cannot disambiguate
 * between split halves).
 */
function getLineFootnoteRefs(
  block: ParagraphBlock,
  fromRun: number,
  toRun: number,
  fnHeights: Map<number, number> | undefined,
): { ids: number[]; height: number } {
  if (!fnHeights) {
    return { ids: [], height: 0 };
  }
  const ids: number[] = [];
  let height = 0;
  for (let r = fromRun; r <= toRun; r++) {
    const run = block.runs[r];
    if (!run || run.kind !== "text") {
      continue;
    }
    const id = (run as { footnoteRefId?: number }).footnoteRefId;
    if (id === undefined) {
      continue;
    }
    const h = fnHeights.get(id);
    if (h !== undefined) {
      ids.push(id);
      height += h;
    }
  }
  return { ids, height };
}

/**
 * Layout a paragraph block onto pages.
 *
 * When `footnoteHeightById` is provided, each line carrying a footnote
 * ref additionally reserves space on its host page for the fn's
 * content. Pagination decisions look at the line's effective height
 * (`lineHeight + lineFnHeight`) so a line cannot land on a page that
 * lacks room for both — preventing footnote overflow without the
 * static-reservation iteration loop.
 */
function layoutParagraph(
  block: ParagraphBlock,
  measure: ParagraphMeasure,
  paginator: ReturnType<typeof createPaginator>,
  contentWidth: number,
  footnoteHeightById?: Map<number, number>,
): void {
  const lines = measure.lines;
  if (lines.length === 0) {
    // Empty paragraph - still takes up space based on spacing
    const spaceBefore = getParagraphSpacingBefore(block);
    const spaceAfter = getParagraphSpacingAfter(block);
    const state = paginator.getCurrentState();

    // Create minimal fragment
    const fragment: ParagraphFragment = {
      kind: "paragraph",
      blockId: block.id,
      x: paginator.getColumnX(state.columnIndex),
      y: state.cursorY + spaceBefore,
      width: contentWidth,
      height: 0,
      fromLine: 0,
      toLine: 0,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
      ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
    };

    paginator.addFragment(fragment, 0, spaceBefore, spaceAfter);
    return;
  }

  const spaceBefore = getParagraphSpacingBefore(block);
  const spaceAfter = getParagraphSpacingAfter(block);

  // Try to fit all lines on current page/column
  let currentLineIndex = 0;

  while (currentLineIndex < lines.length) {
    const state = paginator.getCurrentState();
    const availableHeight = paginator.getAvailableHeight();

    // If the paragraph cannot begin on this page solely because its leading
    // spacing collapses with the previous block's trailing spacing — spacing
    // the next page sheds — advance first so the whole paragraph re-fits there.
    // Otherwise the `fittingLines === 0` fallback below strands its first line
    // here and `addFragment` carries it to the next page alone, splitting a
    // paragraph that fits whole into an artificial intra-page continuation
    // (eigenpal/docx-editor#782 follow-up).
    //
    // Only when the page already holds content (`cursorY !== topMargin`):
    // advancing then lands on a fresh page that resets `trailingSpacing` to 0,
    // so the re-entry takes the normal path. At a fresh page top, `ensureFits`
    // deliberately does not advance oversized content, so continuing there
    // would loop forever — fall through and let the fallback place the line.
    if (currentLineIndex === 0 && state.trailingSpacing > 0 && state.cursorY !== state.topMargin) {
      const firstLine = lines[0]!; // SAFETY: lines.length > 0 in while guard
      const firstLineRefs = getLineFootnoteRefs(
        block,
        firstLine.fromRun,
        firstLine.toRun,
        footnoteHeightById,
      );
      const firstLineHeight = measuredLineAdvance(firstLine) + firstLineRefs.height;
      const collapsedLead = Math.max(spaceBefore, state.trailingSpacing);
      const columnCapacity = state.contentBottom - state.topMargin;
      if (
        collapsedLead + firstLineHeight > availableHeight &&
        spaceBefore + firstLineHeight <= columnCapacity
      ) {
        paginator.ensureFits(collapsedLead + firstLineHeight);
        continue;
      }
    }

    // Calculate how many lines fit
    let linesHeight = 0;
    let linesFnHeight = 0;
    const linesFnIds: number[] = [];
    let fittingLines = 0;

    // The first fragment of a paragraph eats `spaceBefore` from the
    // available height for *every* line check, not only the first one.
    // Pre-fix the loop checked `linesHeight + lineHeight + spaceBefore`
    // only when `j === currentLineIndex`; subsequent lines compared bare
    // line totals against the full available height. That let the loop
    // claim more lines than would actually fit, then `addFragment` (which
    // correctly reserves collapsed spacing + line height) refused the placement
    // and bumped the *whole* fragment to the next page. Result: page-end
    // paragraphs with multi-line content didn't split — they jumped the
    // page boundary, leaving a chunk of empty space above.
    //
    // `addFragment` collapses `spaceBefore` with the previous block's trailing
    // spacing, so the fit loop must reserve the same amount; otherwise a large
    // preceding `spaceAfter` repeats the same over-count (eigenpal/docx-editor#782).
    const firstFragmentSpaceBefore =
      currentLineIndex === 0 ? Math.max(spaceBefore, state.trailingSpacing) : 0;

    for (let j = currentLineIndex; j < lines.length; j++) {
      const line = lines[j]!; // SAFETY: j < lines.length
      const lineAdvance = measuredLineAdvance(line);
      const lineRefs = getLineFootnoteRefs(block, line.fromRun, line.toRun, footnoteHeightById);
      const totalWithLine = linesHeight + lineAdvance;
      const withSpacing =
        totalWithLine + firstFragmentSpaceBefore + linesFnHeight + lineRefs.height;

      if (withSpacing <= availableHeight || fittingLines === 0) {
        linesHeight = totalWithLine;
        linesFnHeight += lineRefs.height;
        for (const id of lineRefs.ids) {
          linesFnIds.push(id);
        }
        fittingLines++;
      } else {
        break;
      }
    }

    let forceBreakAfterFragment = false;
    if (hasWidowControl(block)) {
      const remainingAfter = lines.length - (currentLineIndex + fittingLines);
      if (fittingLines > 1 && remainingAfter === 1) {
        if (currentLineIndex === 0 && fittingLines === 2 && state.cursorY !== state.topMargin) {
          paginator.forceColumnBreak();
          continue;
        }
        fittingLines -= 1;
        forceBreakAfterFragment = true;
        linesHeight = 0;
        linesFnHeight = 0;
        linesFnIds.length = 0;
        for (let j = currentLineIndex; j < currentLineIndex + fittingLines; j++) {
          const line = lines[j]!; // SAFETY: j is within adjusted fitting range
          linesHeight += measuredLineAdvance(line);
          const lineRefs = getLineFootnoteRefs(block, line.fromRun, line.toRun, footnoteHeightById);
          linesFnHeight += lineRefs.height;
          for (const id of lineRefs.ids) {
            linesFnIds.push(id);
          }
        }
      }
    }

    // Create fragment for these lines
    const isFirstFragment = currentLineIndex === 0;
    const isLastFragment = currentLineIndex + fittingLines >= lines.length;
    const effectiveSpaceBefore = isFirstFragment ? spaceBefore : 0;
    const effectiveSpaceAfter = isLastFragment ? spaceAfter : 0;

    const pmRange = getParagraphFragmentPmRange(
      block,
      measure,
      currentLineIndex,
      currentLineIndex + fittingLines,
    );

    const fragment: ParagraphFragment = {
      kind: "paragraph",
      blockId: block.id,
      x: paginator.getColumnX(state.columnIndex),
      y: 0, // Will be set by addFragment
      width: contentWidth,
      height: linesHeight,
      fromLine: currentLineIndex,
      toLine: currentLineIndex + fittingLines,
      ...(pmRange.pmStart !== undefined ? { pmStart: pmRange.pmStart } : {}),
      ...(pmRange.pmEnd !== undefined ? { pmEnd: pmRange.pmEnd } : {}),
      ...(!isFirstFragment ? { continuesFromPrev: true } : {}),
      ...(!isLastFragment ? { continuesOnNext: true } : {}),
      ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
    };

    // Ensure the page can accommodate body lines + footnote demand
    // *together* before placing. Without this, the `fittingLines === 0`
    // fallback in the loop above may force a line carrying a fn ref
    // onto a page that has only a hair of body space left — body fits
    // by itself but `addFootnoteHeight` afterwards drops contentBottom
    // below cursorY, producing an overlap with the fn area.
    //
    // Two-phase check (Codex PR #258 reviews — both edges):
    //
    // 1. Try without the separator overhead. The current page may
    //    already host a footnote — in that case `addFootnoteHeight`
    //    will *not* reserve another separator, so adding 13 px here
    //    would force an unnecessary page advance and split the
    //    paragraph between pages even though the line + fn still fit.
    //
    // 2. After the first `ensureFits`, re-read the page state. If we
    //    landed on a *fresh* page (`footnoteHeight === 0`), the next
    //    `addFootnoteHeight` call *will* reserve a separator — so we
    //    need to verify the line + fn + separator all still fit on
    //    that page. Otherwise the post-commit `addFootnoteHeight`
    //    drops contentBottom past the just-committed line.
    if (linesFnHeight > 0) {
      paginator.ensureFits(effectiveSpaceBefore + linesHeight + linesFnHeight);
      const stateAfter = paginator.getCurrentState();
      if (stateAfter.footnoteHeight === 0) {
        paginator.ensureFits(
          effectiveSpaceBefore + linesHeight + linesFnHeight + FOOTNOTE_SEPARATOR_HEIGHT,
        );
      }
    }

    const result = paginator.addFragment(
      fragment,
      linesHeight,
      effectiveSpaceBefore,
      effectiveSpaceAfter,
    );
    fragment.y = result.y;

    // Now that the lines have committed to this page, grow the page's
    // footnote reservation for any fn refs they carry, and record the
    // IDs on the host page directly. Page → fn-ID mapping is driven by
    // line-level placement here (not by post-layout pmRange mapping)
    // so a fn ref that lives in a continuation fragment of a split
    // paragraph is correctly attributed to the page where the
    // ref-bearing line landed (Codex PR #258 review).
    if (linesFnHeight > 0) {
      paginator.addFootnoteHeight(linesFnHeight, linesFnIds);
    }

    currentLineIndex += fittingLines;

    // If more lines remain, advance to next column/page
    if (currentLineIndex < lines.length) {
      if (forceBreakAfterFragment) {
        paginator.forceColumnBreak();
      } else {
        paginator.ensureFits(measuredLineAdvance(lines[currentLineIndex]!)); // SAFETY: guarded by length check
      }
    }
  }
}

/**
 * Count consecutive header rows at the start of a table.
 * Header rows are marked with isHeader: true in the block data.
 */
function countHeaderRows(block: TableBlock): number {
  let count = 0;
  for (const row of block.rows) {
    if (row.isHeader) {
      count++;
    } else {
      break;
    }
  }
  return count;
}

/**
 * Calculate total height of header rows from their measures.
 */
export function getHeaderRowsHeight(measure: TableMeasure, headerRowCount: number): number {
  let height = 0;
  for (let i = 0; i < headerRowCount && i < measure.rows.length; i++) {
    height += measure.rows[i]!.height; // SAFETY: i < measure.rows.length
  }
  return height;
}

/**
 * Layout a table block onto pages.
 */
function layoutTable(
  block: TableBlock,
  measure: TableMeasure,
  paginator: ReturnType<typeof createPaginator>,
  footnoteHeightById?: Map<number, number>,
): void {
  const rows = measure.rows;
  if (rows.length === 0) {
    return;
  }
  const rowFootnoteIds = block.rows.map((row) =>
    collectTableRowFootnoteIds(row, footnoteHeightById),
  );

  // Detect header rows (consecutive rows at start with isHeader: true)
  const headerRowCount = countHeaderRows(block);
  const headerRowsHeight = getHeaderRowsHeight(measure, headerRowCount);

  let currentRowIndex = 0;

  const breakInfo = buildTableRowBreakInfo(block, measure);

  // X position from justification / indent, recomputed per fragment because the
  // active column can change across section breaks.
  const computeTableX = (columnIndex: number): number => {
    let x = paginator.getColumnX(columnIndex);
    if (block.justification === "center") {
      x += (paginator.columnWidth - measure.totalWidth) / 2;
    } else if (block.justification === "right") {
      x = x + paginator.columnWidth - measure.totalWidth;
    } else if (block.indent !== undefined) {
      // An authored w:tblInd offsets the table border from the content margin.
      // Keep the absent-value path separate because Word's inherited default
      // aligns the first-cell text edge instead.
      x += block.indent;
    } else {
      const leadingCellMargin = block.rows.at(0)?.cells.at(0)?.padding?.left ?? 0;
      x -= leadingCellMargin;
    }
    return x;
  };

  const getCurrentRowCapacity = (state = paginator.getCurrentState()): number =>
    state.rawContentBottom - state.topMargin;

  const getCurrentAvailableHeight = (state = paginator.getCurrentState()): number =>
    state.contentBottom - state.cursorY;

  const hasAdjacentPriorTableRows = (
    rowIndex: number,
    state = paginator.getCurrentState(),
  ): boolean => {
    const previous = state.page.fragments.at(-1);
    return (
      previous?.kind === "table" &&
      previous.blockId === block.id &&
      previous.toRow === rowIndex &&
      previous.y + previous.height === state.cursorY &&
      previous.x === computeTableX(state.columnIndex)
    );
  };

  const shouldRepeatHeaderRows = (
    rowIndex: number,
    consumed: number,
    state = paginator.getCurrentState(),
  ): boolean =>
    headerRowCount > 0 &&
    rowIndex >= headerRowCount &&
    !(consumed === 0 && hasAdjacentPriorTableRows(rowIndex, state));

  const canSplitRow = (rowIndex: number, state = paginator.getCurrentState()): boolean => {
    const row = rows[rowIndex];
    if (!row) {
      return false;
    }
    const repeatedHeaderOverhead = shouldRepeatHeaderRows(rowIndex, 0, state)
      ? headerRowsHeight
      : 0;
    if ((breakInfo.breakOffsets[rowIndex]?.length ?? 0) <= 1) {
      return false;
    }
    const requiredHeight = row.height + repeatedHeaderOverhead;
    if (requiredHeight > getCurrentRowCapacity(state)) {
      return true;
    }
    return (
      hasAdjacentPriorTableRows(rowIndex, state) &&
      requiredHeight > getCurrentAvailableHeight(state) - state.trailingSpacing
    );
  };

  while (currentRowIndex < rows.length) {
    // A row taller than a whole page can't fit anywhere; break it between whole
    // text lines across pages instead of overflowing the page and clipping the
    // content. eigenpal/docx-editor#698 (fixes their #570).
    const oversizedRow = rows[currentRowIndex]!; // SAFETY: currentRowIndex < rows.length
    if (canSplitRow(currentRowIndex)) {
      let consumed = 0;
      while (consumed < oversizedRow.height) {
        const sliceState = paginator.getCurrentState();
        const repeatHeaderRows = shouldRepeatHeaderRows(currentRowIndex, consumed, sliceState);
        const headerOverhead = repeatHeaderRows ? headerRowsHeight : 0;
        const sliceAvail =
          paginator.getAvailableHeight() -
          headerOverhead -
          (consumed === 0 ? sliceState.trailingSpacing : 0);
        let slice = snapRowBreak(breakInfo, currentRowIndex, consumed, sliceAvail);
        if (slice <= 0) {
          const isFreshPage =
            sliceState.cursorY === sliceState.topMargin && sliceState.page.fragments.length === 0;
          if (!isFreshPage) {
            // Not even one line fits in the space left; continue in the next
            // column, or on a fresh page when this is the last column.
            paginator.forceColumnBreak();
            continue;
          }
          // Fresh page and a single line still exceeds the page height: place
          // the next whole line anyway so the loop always makes progress.
          const from = consumed;
          const next = breakInfo.breakOffsets[currentRowIndex]?.find((o) => o > from);
          slice = (next ?? oversizedRow.height) - consumed;
        }
        const sliceBottom = consumed + slice;
        const reachesRowEnd = sliceBottom >= oversizedRow.height;
        const moreAfter = !reachesRowEnd || currentRowIndex + 1 < rows.length;
        const fragmentHeight = headerOverhead + slice;
        const sliceFragment: TableFragment = {
          kind: "table",
          blockId: block.id,
          x: computeTableX(sliceState.columnIndex),
          y: 0,
          width: measure.totalWidth,
          height: fragmentHeight,
          fromRow: currentRowIndex,
          toRow: currentRowIndex + 1,
          ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
          ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
          ...(consumed > 0 || currentRowIndex > 0 ? { continuesFromPrev: true } : {}),
          ...(moreAfter ? { continuesOnNext: true } : {}),
          ...(repeatHeaderRows ? { headerRowCount } : {}),
          ...(consumed > 0 ? { topClip: consumed } : {}),
          ...(reachesRowEnd ? {} : { bottomClip: sliceBottom }),
          ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
        };
        const sliceResult = paginator.addFragment(sliceFragment, fragmentHeight, 0, 0);
        sliceFragment.y = sliceResult.y;
        sliceFragment.x = computeTableX(sliceResult.state.columnIndex);
        consumed = sliceBottom;
        if (consumed < oversizedRow.height) {
          paginator.forceColumnBreak();
        }
      }
      currentRowIndex += 1;
      continue;
    }

    const state = paginator.getCurrentState();
    const rawAvailableHeight = paginator.getAvailableHeight();
    const isFirstFragment = currentRowIndex === 0;

    // Leading skip past a page-pinned band, applied only to the table's first
    // fragment (a band sits on one page). eigenpal #694.
    const bandSkip = isFirstFragment ? (measure.bandSkipBefore ?? 0) : 0;

    // Account for the space addFragment will consume before the fragment, which
    // is max(spaceBefore, trailingSpacing). We pass bandSkip as spaceBefore, so
    // the overhead is the larger of that and the previous block's trailing space.
    const pendingSpacing = isFirstFragment ? Math.max(bandSkip, state.trailingSpacing) : 0;
    const availableHeight = rawAvailableHeight - pendingSpacing;

    const repeatHeaderRowsForNormalFragment = shouldRepeatHeaderRows(currentRowIndex, 0, state);

    // For continuation fragments, we need space for header rows + at least one content row.
    const normalHeaderOverhead = repeatHeaderRowsForNormalFragment ? headerRowsHeight : 0;

    // Calculate how many rows fit (excluding header rows which are prepended separately)
    let rowsHeight = 0;
    let fittingRows = 0;
    const pageFootnoteIds = new Set(state.page.footnoteIds ?? []);
    const fragmentFootnoteIds: number[] = [];
    let fragmentFootnoteHeight = 0;
    let retryOnNextFlowRegion = false;

    for (let j = currentRowIndex; j < rows.length; j++) {
      const rowHeight = rows[j]!.height; // SAFETY: j < rows.length
      const currentRowFootnoteIds = rowFootnoteIds[j] ?? [];
      const fragmentFootnoteCountBeforeRow = fragmentFootnoteIds.length;
      let rowFootnoteHeight = 0;
      for (const id of currentRowFootnoteIds) {
        if (pageFootnoteIds.has(id) || fragmentFootnoteIds.includes(id)) {
          continue;
        }
        fragmentFootnoteIds.push(id);
        rowFootnoteHeight += footnoteHeightById?.get(id) ?? 0;
      }
      const separatorHeight =
        pageFootnoteIds.size === 0 && fragmentFootnoteHeight + rowFootnoteHeight > 0
          ? FOOTNOTE_SEPARATOR_HEIGHT
          : 0;
      const totalWithRow =
        rowsHeight +
        rowHeight +
        normalHeaderOverhead +
        fragmentFootnoteHeight +
        rowFootnoteHeight +
        separatorHeight;

      if (totalWithRow <= availableHeight) {
        rowsHeight += rowHeight;
        fragmentFootnoteHeight += rowFootnoteHeight;
        fittingRows++;
      } else if (fittingRows === 0) {
        const isFreshFlowRegion =
          state.cursorY === state.topMargin && state.page.fragments.length === 0;
        if (isFreshFlowRegion) {
          rowsHeight += rowHeight;
          fragmentFootnoteHeight += rowFootnoteHeight;
          fittingRows++;
          break;
        }
        fragmentFootnoteIds.splice(fragmentFootnoteCountBeforeRow);
        paginator.forceColumnBreak();
        retryOnNextFlowRegion = true;
        break;
      } else {
        fragmentFootnoteIds.splice(fragmentFootnoteCountBeforeRow);
        break;
      }
    }

    if (retryOnNextFlowRegion) {
      continue;
    }

    // Total fragment height includes header rows for continuation fragments
    const fragmentHeight = rowsHeight + normalHeaderOverhead;

    // Create fragment for these rows
    const isLastFragment = currentRowIndex + fittingRows >= rows.length;

    // Calculate x position based on table justification and indent
    const desiredX = computeTableX(state.columnIndex);

    const fragment: TableFragment = {
      kind: "table",
      blockId: block.id,
      x: desiredX,
      y: 0, // Will be set by addFragment
      width: measure.totalWidth,
      height: fragmentHeight,
      fromRow: currentRowIndex,
      toRow: currentRowIndex + fittingRows,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
      ...(!isFirstFragment ? { continuesFromPrev: true } : {}),
      ...(!isLastFragment ? { continuesOnNext: true } : {}),
      ...(repeatHeaderRowsForNormalFragment ? { headerRowCount } : {}),
      ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
    };

    if (fragmentFootnoteHeight > 0) {
      paginator.addFootnoteHeight(fragmentFootnoteHeight, fragmentFootnoteIds);
    }
    const result = paginator.addFragment(fragment, fragmentHeight, bandSkip, 0);
    fragment.y = result.y;
    fragment.x = desiredX;

    currentRowIndex += fittingRows;

    // If more rows remain, advance to next column/page
    if (currentRowIndex < rows.length) {
      if (canSplitRow(currentRowIndex)) {
        const nextState = paginator.getCurrentState();
        const nextHeaderOverhead = shouldRepeatHeaderRows(currentRowIndex, 0, nextState)
          ? headerRowsHeight
          : 0;
        const nextSliceAvail =
          paginator.getAvailableHeight() - nextHeaderOverhead - nextState.trailingSpacing;
        if (snapRowBreak(breakInfo, currentRowIndex, 0, nextSliceAvail) > 0) {
          continue;
        }
      }
      // Need space for at least one content row plus repeated header rows
      const nextState = paginator.getCurrentState();
      const nextRowHeight =
        rows[currentRowIndex]!.height + // SAFETY: guarded by length check
        (shouldRepeatHeaderRows(currentRowIndex, 0, nextState) ? headerRowsHeight : 0);
      paginator.ensureFits(nextRowHeight);
    }
  }
}

function collectTableRowFootnoteIds(
  row: TableBlock["rows"][number] | undefined,
  footnoteHeightById: Map<number, number> | undefined,
): number[] {
  if (!row || !footnoteHeightById) {
    return [];
  }

  const ids: number[] = [];
  const walk = (blocks: FlowBlock[]): void => {
    for (const block of blocks) {
      if (block.kind === "paragraph") {
        for (const run of block.runs) {
          if (
            run.kind === "text" &&
            run.footnoteRefId !== undefined &&
            footnoteHeightById.has(run.footnoteRefId) &&
            !ids.includes(run.footnoteRefId)
          ) {
            ids.push(run.footnoteRefId);
          }
        }
        continue;
      }
      if (block.kind === "table") {
        for (const nestedRow of block.rows) {
          for (const cell of nestedRow.cells) {
            walk(cell.blocks);
          }
        }
        continue;
      }
      if (block.kind === "textBox") {
        walk(block.content);
      }
    }
  };

  for (const cell of row.cells) {
    walk(cell.blocks);
  }
  return ids;
}

/**
 * Layout a floating table (anchored) without advancing the cursor.
 */
function layoutFloatingTable(
  block: TableBlock,
  measure: TableMeasure,
  paginator: ReturnType<typeof createPaginator>,
  contentWidth: number,
): void {
  const state = paginator.getCurrentState();
  const floating = block.floating;
  const page = state.page;
  const margins = page.margins;

  const tableWidth = measure.totalWidth;
  const tableHeight = measure.totalHeight;

  const contentHeight = page.size.h - margins.top - margins.bottom;

  // Default anchor base (content area)
  let baseX = margins.left;
  let baseY = margins.top;

  if (floating?.horzAnchor === "page") {
    baseX = 0;
  }
  if (floating?.vertAnchor === "page") {
    baseY = 0;
  } else if (floating?.vertAnchor === "text") {
    // A text-relative table position is offset from the current text
    // anchor, not from the page's body margin. Using the margin here lifts
    // later floating tables to the top of the page and overlays earlier text.
    baseY = state.cursorY;
  }

  // Determine X position
  let x = paginator.getColumnX(state.columnIndex);
  if (floating?.tblpX !== undefined) {
    x = baseX + floating.tblpX;
  } else if (floating?.tblpXSpec) {
    const spec = floating.tblpXSpec;
    if (spec === "left" || spec === "inside") {
      x = baseX;
    } else if (spec === "right" || spec === "outside") {
      x = baseX + contentWidth - tableWidth;
    } else {
      x = baseX + (contentWidth - tableWidth) / 2;
    }
  } else if (block.justification === "center") {
    x = baseX + (contentWidth - tableWidth) / 2;
  } else if (block.justification === "right") {
    x = baseX + contentWidth - tableWidth;
  }

  // Determine Y position
  let y = state.cursorY;
  let usedExplicitY = false;
  if (floating?.tblpY !== undefined) {
    y = baseY + floating.tblpY;
    usedExplicitY = true;
  } else if (floating?.tblpYSpec) {
    usedExplicitY = true;
    const spec = floating.tblpYSpec;
    if (spec === "top") {
      y = baseY;
    } else if (spec === "bottom") {
      y = baseY + contentHeight - tableHeight;
    } else if (spec === "center") {
      y = baseY + (contentHeight - tableHeight) / 2;
    }
  }

  // If not explicitly positioned, ensure it fits on the current page
  if (!usedExplicitY) {
    const fitState = paginator.ensureFits(tableHeight);
    y = fitState.cursorY;
  }

  // Alignment keywords stay inside their selected anchor frame. A numeric
  // offset may deliberately move a margin-anchored table into the page margin,
  // so clamp that resolved position only against the physical page.
  const usesNumericOffset = floating?.tblpX !== undefined;
  const pageAnchored = floating?.horzAnchor === "page";
  const clampToPage = pageAnchored || usesNumericOffset;
  const minX = clampToPage ? 0 : margins.left;
  const maxX = clampToPage ? page.size.w - tableWidth : margins.left + contentWidth - tableWidth;
  if (Number.isFinite(maxX)) {
    x = Math.max(minX, Math.min(x, maxX));
  }

  const fragment: TableFragment = {
    kind: "table",
    blockId: block.id,
    x,
    y,
    width: tableWidth,
    height: tableHeight,
    fromRow: 0,
    toRow: block.rows.length,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    isFloating: true,
    ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
  };

  // Add directly without advancing cursor
  state.page.fragments.push(fragment);
}

/**
 * Layout an image block onto pages.
 */
function layoutImage(
  block: ImageBlock,
  measure: ImageMeasure,
  paginator: ReturnType<typeof createPaginator>,
): void {
  // Handle anchored images differently
  if (block.anchor?.isAnchored) {
    layoutAnchoredImage(block, measure, paginator);
    return;
  }

  // Inline image - ensure it fits (plus any leading skip past a page band)
  const bandSkip = measure.bandSkipBefore ?? 0;
  const state = paginator.ensureFits(bandSkip + measure.height);

  const fragment: ImageFragment = {
    kind: "image",
    blockId: block.id,
    x: paginator.getColumnX(state.columnIndex),
    y: 0, // Will be set by addFragment
    width: measure.width,
    height: measure.height,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
  };

  const result = paginator.addFragment(fragment, measure.height, bandSkip, 0);
  fragment.y = result.y;
}

/**
 * Layout an anchored (floating) image.
 */
function layoutAnchoredImage(
  block: ImageBlock,
  measure: ImageMeasure,
  paginator: ReturnType<typeof createPaginator>,
): void {
  const state = paginator.getCurrentState();
  const anchor = block.anchor;
  if (!anchor) {
    return;
  }

  // Position based on anchor offsets
  const x = anchor.offsetH ?? paginator.getColumnX(state.columnIndex);
  const y = anchor.offsetV ?? state.cursorY;

  const fragment: ImageFragment = {
    kind: "image",
    blockId: block.id,
    x,
    y,
    width: measure.width,
    height: measure.height,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    isAnchored: true,
    zIndex: anchor.behindDoc ? -1 : 1,
  };

  // Add directly to page without affecting cursor
  state.page.fragments.push(fragment);
}

type LayoutTextBoxOptions = {
  paginator: ReturnType<typeof createPaginator>;
  sectionMarginTop: number;
  sectionPageHeight: number;
  sectionMarginBottom: number;
};

/**
 * Layout a text box block onto pages.
 */
function layoutTextBox(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  { paginator, sectionMarginTop, sectionPageHeight, sectionMarginBottom }: LayoutTextBoxOptions,
): void {
  // A page/margin-pinned topAndBottom band (e.g. a title banner) floats to the
  // top of its page; the reserved band in the measure pass pushes body text
  // below it (see extractFloatingZones in PagedEditor). It must not also consume
  // flow at its anchor, or the box height would be reserved twice. Place it at
  // the page content top without advancing the cursor. eigenpal #694.
  if (isPagePinnedBandTextBox(block)) {
    const state = paginator.getCurrentState();
    // Position the box at the same content-Y the measure pass reserved its band
    // at. The measure pass uses the section top margin (not a page's
    // first-page margin), so use the same value here; `bandTopContentY` is
    // content-relative, and `fragment.y` is page-absolute, so add the page's
    // own top margin (`state.topMargin`) to convert. Using `state.topMargin`
    // inside the resolver instead would desync the box from its band on a
    // title page whose first-page top margin differs from the section margin.
    const bandTop = bandTopContentY(block.position?.vertical, {
      pageHeight: sectionPageHeight,
      marginTop: sectionMarginTop,
      marginBottom: sectionMarginBottom,
      boxHeight: measure.height,
    });
    // Honor the box's horizontal anchor (align center/right, page-relative
    // offset) instead of always pinning to the column's left edge. The band is
    // full-width regardless, so this only moves where the box paints.
    const horizontal = block.position?.horizontal;
    const x = horizontal
      ? bandFragmentX(horizontal, {
          pageWidth: state.page.size.w,
          marginLeft: state.page.margins.left,
          marginRight: state.page.margins.right,
          boxWidth: measure.width,
        })
      : paginator.getColumnX(state.columnIndex);
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: block.id,
      x,
      y: state.topMargin + bandTop,
      width: measure.width,
      height: measure.height,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    };
    state.page.fragments.push(fragment);
    return;
  }

  // Any explicitly positioned textbox is an anchored object, including
  // paragraph/line-relative anchors. Place it from the current text anchor
  // without advancing normal flow; otherwise several shapes owned by one
  // shape-only paragraph stack vertically and push the body onto another page.
  if (block.position !== undefined) {
    const state = paginator.getCurrentState();
    const horizontal = block.position.horizontal;
    const x = horizontal
      ? bandFragmentX(horizontal, {
          pageWidth: state.page.size.w,
          marginLeft: state.page.margins.left,
          marginRight: state.page.margins.right,
          boxWidth: measure.width,
        })
      : paginator.getColumnX(state.columnIndex);
    const vertical = block.position.vertical;
    const y = isPageFrameRelativeAnchor(vertical?.relativeTo)
      ? state.topMargin +
        bandTopContentY(vertical, {
          pageHeight: sectionPageHeight,
          marginTop: sectionMarginTop,
          marginBottom: sectionMarginBottom,
          boxHeight: measure.height,
        })
      : state.cursorY + emuToPixels(vertical?.posOffset ?? 0);
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: block.id,
      x,
      y,
      width: measure.width,
      height: measure.height,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    };
    state.page.fragments.push(fragment);
    return;
  }

  const state = paginator.ensureFits(measure.height);

  const fragment: TextBoxFragment = {
    kind: "textBox",
    blockId: block.id,
    x: paginator.getColumnX(state.columnIndex),
    y: 0,
    width: measure.width,
    height: measure.height,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
  };

  const result = paginator.addFragment(fragment, measure.height, 0, 0);
  fragment.y = result.y;
}

/**
 * A topAndBottom text box whose vertical anchor pins it to the page frame
 * (page/margin/margin-strip) — it floats to a fixed page position rather than
 * flowing in document order. Must agree with the measure pass's band extraction
 * (extractFloatingZones), which uses the same predicate. eigenpal #694.
 */
function isPagePinnedBandTextBox(block: TextBoxBlock): boolean {
  if (!floatingTextBoxReservesBand(block)) {
    return false;
  }
  return isPageFrameRelativeAnchor(block.position?.vertical?.relativeTo);
}

/**
 * Handle a section break block.
 * @param block - The section break block (current section's properties)
 * @param paginator - The paginator instance
 * @param nextSectionConfig - Page layout for the NEXT section
 * @param nextSectionType - Break type of the NEXT section (how it starts relative to current)
 */
function handleSectionBreak(
  _block: SectionBreakBlock,
  paginator: ReturnType<typeof createPaginator>,
  nextSectionConfig: SectionLayoutConfig,
  nextSectionType: SectionBreakBlock["type"] = "nextPage",
  nextSectionIndex?: number,
): void {
  switch (nextSectionType) {
    case "nextPage":
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins);
      if (nextSectionIndex !== undefined) {
        paginator.startSection(nextSectionIndex);
      }
      paginator.forcePageBreak({ coalesceBlankPage: true });
      break;

    case "evenPage": {
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins);
      if (nextSectionIndex !== undefined) {
        paginator.startSection(nextSectionIndex);
      }
      const state = paginator.forcePageBreak({ coalesceBlankPage: true });
      // If landed on odd page, add another page
      if (state.page.number % 2 !== 0) {
        paginator.forcePageBreak();
      }
      break;
    }

    case "oddPage": {
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins);
      if (nextSectionIndex !== undefined) {
        paginator.startSection(nextSectionIndex);
      }
      const state = paginator.forcePageBreak({ coalesceBlankPage: true });
      // If landed on even page, add another page
      if (state.page.number % 2 === 0) {
        paginator.forcePageBreak();
      }
      break;
    }

    case "continuous": {
      // ECMA-376 §17.6.22: a `continuous` break normally keeps the current
      // page geometry and defers the new size/margins to the next natural
      // page break. But a break that changes page size or orientation cannot
      // share a physical sheet with the preceding section, so Word and
      // LibreOffice promote it to a page break (eigenpal/docx-editor#841).
      //
      // Compare against the last laid-out page without materializing one: a
      // break before any content has no sheet to share, so it defers (the first
      // content then opens a page with the new geometry) rather than stranding
      // a blank leading page.
      const currentPage = paginator.pages.at(-1);
      const nextSize = nextSectionConfig.pageSize;
      const pageSizeChanges =
        currentPage != null &&
        (Math.round(nextSize.w) !== Math.round(currentPage.size.w) ||
          Math.round(nextSize.h) !== Math.round(currentPage.size.h));
      if (nextSectionIndex !== undefined) {
        paginator.startSection(nextSectionIndex);
      }
      if (pageSizeChanges) {
        // Promote to a page break, but reuse an already blank current page as
        // the next section's first page instead of leaving it stranded.
        paginator.updatePageLayout(nextSize, nextSectionConfig.margins);
        if (!paginator.retargetCurrentBlankPage()) {
          paginator.forcePageBreak({ coalesceBlankPage: true });
        }
      } else {
        paginator.updatePageLayout(nextSize, nextSectionConfig.margins, false);
        paginator.retargetCurrentBlankPage();
      }
      break;
    }
    default:
      break;
  }

  // Update column layout for the next section
  paginator.updateColumns(nextSectionConfig.columns ?? DEFAULT_COLUMNS);
}

// Re-export types
export * from "./types";
export { createPaginator } from "./paginator";
export type { PageState, PaginatorOptions, Paginator } from "./paginator";
export {
  computeKeepNextChains,
  calculateChainHeight,
  getMidChainIndices,
  hasKeepLines,
  hasPageBreakBefore,
} from "./keep-together";
export type { KeepNextChain } from "./keep-together";
export {
  scheduleSectionBreak,
  applyPendingToActive,
  createInitialSectionState,
  getEffectiveMargins,
  getEffectivePageSize,
  getEffectiveColumns,
} from "./section-breaks";
export type { SectionState, BreakDecision } from "./section-breaks";
export { assertExhaustiveFlowBlock, findPageIndexContainingPmPos } from "./pmPageIndex";
