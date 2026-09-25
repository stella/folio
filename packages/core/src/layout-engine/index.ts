/**
 * Layout Engine - Main Entry Point
 *
 * Converts blocks + measures into positioned fragments on pages.
 */

import { panic } from "better-result";
import { reflowFootnoteColumns } from "./footnoteColumnReflow";
import {
  computeKeepNextChains,
  calculateChainHeight,
  getMidChainIndices,
  hasPageBreakBefore,
} from "./keep-together";
import { measuredLineAdvance } from "./lineFlow";
import {
  annotateNoteSeparators,
  continuesNoteArea,
  createNoteAreaContinuation,
} from "./noteAreaFlow";
import { createPaginator } from "./paginator";
import { paragraphsShareStyle, resolveEffectiveParagraphSpacingTree } from "./paragraphSpacing";
import {
  INITIAL_RENDERED_BREAK_STATE,
  PAGE_ADVANCE,
  reconcileAfterBlock,
  reconcileBreakBeforeBlock,
  recordReflowBoundary,
} from "./renderedBreakReconciliation";
import { normalizeSectionBreakType } from "./section-breaks";
import { applySectionVerticalAlignment } from "./sectionVerticalAlignment";
import { tableKeepNextOpeningHeight } from "./tableRowBreak";
import type {
  FlowBlock,
  Measure,
  Layout,
  LayoutOptions,
  PageMargins,
  ParagraphMeasure,
  TableMeasure,
  ImageMeasure,
  TextBoxBlock,
  TextBoxMeasure,
  SectionBreakBlock,
} from "./types";
import { RENDERED_BREAK_REFLOW_TOLERANCE_LINES } from "./layoutFlowShared";
import {
  DEFAULT_COLUMNS,
  CONTINUE_PAGE_NUMBERING,
  collectSectionConfigs,
  balancedParagraphSectionHeight,
  handleSectionBreak,
} from "./sectionLayout";
import type { SectionLayoutConfig } from "./sectionLayout";
import { layoutParagraph } from "./paragraphLayout";
import { layoutTable, layoutFloatingTable } from "./tableLayout";
import { layoutImage } from "./imageLayout";
import { layoutTextBox } from "./textBoxLayout";

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
export { resolveSectionHeaderFooterRefs } from "./headerFooterRefs";
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
export { collectSectionConfigs } from "./sectionLayout";
export type { SectionLayoutConfig } from "./sectionLayout";
export { getHeaderRowsHeight } from "./tableLayout";

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

    if (!paragraphsShareStyle(curr, next)) {
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

/**
 * Layout a document: convert blocks + measures into pages with positioned fragments.
 *
 * A reference discovered after the first column has already been placed can shrink the
 * shared body band below earlier-column fragments. Retry those documents with the observed
 * reservations as page floors so every column participates in the same footnote geometry.
 */
export function layoutDocument(
  blocks: FlowBlock[],
  measures: Measure[],
  options: LayoutOptions,
): Layout {
  const initialLayout = layoutDocumentPass(blocks, measures, options);
  const layout = options.footnoteHeightById
    ? reflowFootnoteColumns({
        initialLayout,
        ...(options.footnoteReservedHeights
          ? { initialReserveFloors: options.footnoteReservedHeights }
          : {}),
        runLayout: (reserveFloors) =>
          layoutDocumentPass(blocks, measures, {
            ...options,
            footnoteReservedHeights: reserveFloors,
          }),
      })
    : initialLayout;

  return annotateNoteSeparators(
    applySectionVerticalAlignment(layout, options.sectionVerticalAlignments),
    blocks,
    measures,
    options.noteAreas,
  );
}

function layoutDocumentPass(
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
    pageNumbering: options.pageNumbering ?? CONTINUE_PAGE_NUMBERING,
  };
  if (options.columns !== undefined) {
    bodyConfig.columns = options.columns;
  }
  const finalConfig: SectionLayoutConfig = {
    pageSize: finalPageSize,
    margins: finalMargins,
    pageNumbering: options.finalPageNumbering ?? bodyConfig.pageNumbering,
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
  const sectionBreakTypes = breakIndices.map((index) => (blocks[index] as SectionBreakBlock).type);
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
    ...(options.sectionFirstPageMargins !== undefined
      ? { sectionFirstPageMargins: options.sectionFirstPageMargins }
      : {}),
    columns: initialConfig.columns ?? DEFAULT_COLUMNS,
    pageNumbering: initialConfig.pageNumbering,
    ...(options.footnoteReservedHeights !== undefined
      ? { footnoteReservedHeights: options.footnoteReservedHeights }
      : {}),
    ...(options.sectionHeaderFooterRefs !== undefined
      ? { sectionHeaderFooterRefs: options.sectionHeaderFooterRefs }
      : {}),
  });

  // Resolve contextual and automatic-list spacing without mutating the input
  // flow tree. Re-resolution is idempotent for pipeline-prepared blocks.
  blocks = resolveEffectiveParagraphSpacingTree(blocks);

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
  let renderedBreakState = INITIAL_RENDERED_BREAK_STATE;
  const noteAreaContinuation = createNoteAreaContinuation(options.noteAreas, paginator);
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]!; // SAFETY: i < blocks.length
    const measure = measures[i]!; // SAFETY: measures.length === blocks.length (validated above)
    if (noteAreaContinuation !== undefined) {
      paginator.setPageTopContinuation(
        continuesNoteArea(options.noteAreas, block) ? noteAreaContinuation : undefined,
      );
    }

    const firstLine = measure.kind === "paragraph" ? measure.lines.at(0) : undefined;
    const firstLineAdvance = firstLine ? measuredLineAdvance(firstLine) : 0;
    const renderedBreakNeedsSnap =
      measure.kind === "paragraph" &&
      (!paginator.fits(measure.totalHeight) ||
        (firstLineAdvance > 0 &&
          paginator.getAvailableHeight() <=
            firstLineAdvance * RENDERED_BREAK_REFLOW_TOLERANCE_LINES));
    const hasExplicitPageBreak = block.kind === "pageBreak" || hasPageBreakBefore(block);
    const breakDecision = reconcileBreakBeforeBlock({
      state: renderedBreakState,
      block,
      previousBlock: blocks[i - 1],
      page: paginator.getCurrentState().page,
      blocksById,
      hasExplicitPageBreak,
      renderedBreakNeedsSnap,
    });
    if (block.kind !== "pageBreak") {
      if (breakDecision.pageAdvance === PAGE_ADVANCE.PHYSICAL) {
        // w:pageBreakBefore puts the paragraph at the top of a page. A page
        // that holds nothing yet but zero-height carriers (such as the mark
        // of a paragraph that ended in a page break) already satisfies that,
        // so the paragraph starts there instead of leaving it blank.
        paginator.forcePageBreak({ coalesceBlankPage: hasPageBreakBefore(block) });
      } else if (breakDecision.pageAdvance === PAGE_ADVANCE.COALESCED) {
        paginator.coalescePageBreak();
      }
    }
    renderedBreakState = breakDecision.state;

    // Handle keepNext chains - if this is a chain start, check if chain fits
    const chain = keepNextChains.get(i);
    if (chain && !midChainIndices.has(i)) {
      const chainHeight = calculateChainHeight(
        chain,
        blocks,
        measures,
        paginator.getCurrentState().trailingSpacing,
        tableKeepNextOpeningHeight,
      );
      const pageBeforeChainLayout = paginator.getCurrentState().page.number;
      paginator.ensureFits(chainHeight);
      if (paginator.getCurrentState().page.number > pageBeforeChainLayout) {
        // The chain moved as one unit across the cached boundary. A rendered
        // marker immediately after it describes the page Folio just opened.
        renderedBreakState = recordReflowBoundary(renderedBreakState, true);
      }
    }

    const pageBeforeBlockLayout = paginator.getCurrentState().page.number;
    switch (block.kind) {
      case "paragraph":
        layoutParagraph({
          block,
          measure: measure as ParagraphMeasure,
          paginator,
          contentWidth: paginator.columnWidth,
          footnoteHeightById: options.footnoteHeightById,
          suppressSpaceBefore: breakDecision.suppressSpaceBefore,
        });
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
        if (breakDecision.pageAdvance === PAGE_ADVANCE.PHYSICAL) {
          paginator.forcePageBreak();
        } else if (breakDecision.pageAdvance === PAGE_ADVANCE.COALESCED) {
          paginator.coalescePageBreak();
        }
        break;

      case "columnBreak":
        paginator.forceColumnBreak();
        break;

      case "sectionBreak": {
        const nextSectionConfig = sectionConfigs[sectionIdx + 1] ?? initialConfig;
        const nextType = normalizeSectionBreakType(sectionBreakTypes[sectionIdx]);
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

    renderedBreakState = reconcileAfterBlock({
      state: renderedBreakState,
      block,
      pageNumberBefore: pageBeforeBlockLayout,
      pageNumberAfter: paginator.getCurrentState().page.number,
      previousPage: paginator.states[pageBeforeBlockLayout - 1]?.page,
    });
  }

  paginator.setPageTopContinuation(undefined);

  // Ensure at least one page exists
  if (paginator.states.length === 0) {
    paginator.getCurrentState();
  }

  return {
    pageSize,
    pages: paginator.pages,
    ...(options.columns !== undefined ? { columns: options.columns } : {}),
    ...(options.pageGap !== undefined ? { pageGap: options.pageGap } : {}),
  };
}
