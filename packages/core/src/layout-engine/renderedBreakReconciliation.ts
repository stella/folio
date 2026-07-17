import { continuesNumberedSequence, continuesTabbedParagraphSequence } from "./paragraphSequence";
import type { FlowBlock, Page, ParagraphBlock } from "./types";

export type RenderedBreakState =
  | { type: "noPageAdvance" }
  | { type: "pageAdvance"; reason: "ordinary" | "reflowBoundary" }
  | {
      type: "pageAdvance";
      reason: "authoredBoundary";
      boundary: "hardBreak" | "sectionBreak";
    };

export const INITIAL_RENDERED_BREAK_STATE: RenderedBreakState = {
  type: "noPageAdvance",
};

type ReconcileBeforeBlockOptions = {
  state: RenderedBreakState;
  block: FlowBlock;
  previousBlock: FlowBlock | undefined;
  page: Page;
  blocksById: ReadonlyMap<string, FlowBlock>;
  hasExplicitPageBreak: boolean;
  renderedBreakNeedsSnap: boolean;
};

type BreakDecision = {
  forcePageBreak: boolean;
  suppressSpaceBefore: boolean;
  state: RenderedBreakState;
};

/**
 * Reconcile an authored or cached break before laying out one block.
 *
 * Cached markers describe a boundary at a paragraph, not an ordinal page.
 * Natural reflow may therefore satisfy the marker before layout reaches it.
 * Every marker consumes the accumulated advance state so later markers make
 * independent decisions.
 */
export const reconcileBreakBeforeBlock = ({
  state,
  block,
  previousBlock,
  page,
  blocksById,
  hasExplicitPageBreak,
  renderedBreakNeedsSnap,
}: ReconcileBeforeBlockOptions): BreakDecision => {
  if (hasExplicitPageBreak) {
    return {
      forcePageBreak: true,
      suppressSpaceBefore: false,
      state: INITIAL_RENDERED_BREAK_STATE,
    };
  }
  if (block.kind !== "paragraph" || block.attrs?.renderedPageBreakBefore !== true) {
    return { forcePageBreak: false, suppressSpaceBefore: false, state };
  }

  const markerAlreadySatisfied =
    pageStartsWithPreviousParagraphContinuation(page, previousBlock) ||
    (state.type === "pageAdvance" && state.reason === "authoredBoundary") ||
    (state.type === "pageAdvance" && state.reason === "reflowBoundary") ||
    (state.type === "pageAdvance" && isPaginationEmptyParagraph(block)) ||
    (state.type === "pageAdvance" &&
      (continuesNumberedSequence(previousBlock, block) ||
        continuesTabbedParagraphSequence(previousBlock, block)));
  // A keep-with-next paragraph carries the marker boundary into its linked
  // content, so its own height is not enough to classify the marker as stale.
  const markerNeedsSnap =
    renderedBreakNeedsSnap || block.attrs?.keepNext === true || previousBlock?.kind === "table";
  const forcePageBreak =
    markerNeedsSnap && !markerAlreadySatisfied && pageHasVisibleBodyContent(page, blocksById);
  const followsAuthoredPageBreak = previousBlock?.kind === "pageBreak";
  const followsSectionBreak = previousBlock?.kind === "sectionBreak";
  const followsPageBreakingSection = followsSectionBreak && previousBlock.type !== "continuous";
  const followsCarriedSectionBoundary =
    state.type === "pageAdvance" &&
    state.reason === "authoredBoundary" &&
    state.boundary === "sectionBreak";
  const preserveSectionLeadingSpacing =
    (followsSectionBreak || followsCarriedSectionBoundary) &&
    block.attrs?.spacingExplicit?.before === true;
  const markerAccountsForBoundary =
    forcePageBreak ||
    markerAlreadySatisfied ||
    followsAuthoredPageBreak ||
    followsPageBreakingSection;

  return {
    forcePageBreak,
    // The boundary consumes inherited/default leading spacing. Preserve only
    // spacing authored directly on the first paragraph of a section.
    suppressSpaceBefore: markerAccountsForBoundary && !preserveSectionLeadingSpacing,
    state: INITIAL_RENDERED_BREAK_STATE,
  };
};

export const recordReflowBoundary = (
  state: RenderedBreakState,
  crossedPageBoundary: boolean,
): RenderedBreakState =>
  crossedPageBoundary ? { type: "pageAdvance", reason: "reflowBoundary" } : state;

type ReconcileAfterBlockOptions = {
  state: RenderedBreakState;
  block: FlowBlock;
  pageNumberBefore: number;
  pageNumberAfter: number;
  previousPage: Page | undefined;
};

/** Record page movement and carry authored boundaries across pagination-empty paragraphs. */
export const reconcileAfterBlock = ({
  state,
  block,
  pageNumberBefore,
  pageNumberAfter,
  previousPage,
}: ReconcileAfterBlockOptions): RenderedBreakState => {
  if (isStructuralBreak(block)) {
    if (pageNumberAfter <= pageNumberBefore) {
      if (state.type === "pageAdvance" && state.reason === "authoredBoundary") {
        return block.kind === "sectionBreak" ? { ...state, boundary: "sectionBreak" } : state;
      }
      return INITIAL_RENDERED_BREAK_STATE;
    }
    return {
      type: "pageAdvance",
      reason: "authoredBoundary",
      boundary: block.kind === "sectionBreak" ? "sectionBreak" : "hardBreak",
    };
  }
  if (block.kind === "paragraph" && isPaginationEmptyParagraph(block)) {
    return state;
  }
  if (pageNumberAfter <= pageNumberBefore) {
    return state.type === "pageAdvance" && state.reason === "authoredBoundary"
      ? INITIAL_RENDERED_BREAK_STATE
      : state;
  }

  const blockContinuedAcrossBoundary = previousPage?.fragments.some(
    ({ blockId }) => blockId === block.id,
  );
  return {
    type: "pageAdvance",
    reason: blockContinuedAcrossBoundary ? "reflowBoundary" : "ordinary",
  };
};

const isPaginationEmptyParagraph = (block: ParagraphBlock): boolean =>
  !hasVisibleEmptyParagraphDecoration(block) &&
  block.runs.every((run) => {
    if (run.kind === "text") {
      return run.text.trim().length === 0;
    }
    return run.kind === "tab" || run.kind === "lineBreak";
  });

const hasVisibleEmptyParagraphDecoration = (block: ParagraphBlock): boolean => {
  if (
    block.attrs?.listMarker !== undefined &&
    block.attrs.listMarker.length > 0 &&
    block.attrs.listMarkerHidden !== true
  ) {
    return true;
  }
  if (block.attrs?.shading) {
    return true;
  }
  return Object.values(block.attrs?.borders ?? {}).some(
    (border) =>
      border !== undefined &&
      border.style !== "none" &&
      border.style !== "nil" &&
      border.width !== 0,
  );
};

const pageHasVisibleBodyContent = (
  page: Page,
  blocksById: ReadonlyMap<string, FlowBlock>,
): boolean => {
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
};

const pageStartsWithPreviousParagraphContinuation = (
  page: Page,
  previousBlock: FlowBlock | undefined,
): boolean => {
  if (previousBlock?.kind !== "paragraph") {
    return false;
  }

  return page.fragments.some(
    (fragment) =>
      fragment.kind === "paragraph" &&
      String(fragment.blockId) === String(previousBlock.id) &&
      fragment.continuesFromPrev === true,
  );
};

const isStructuralBreak = (block: FlowBlock): boolean =>
  block.kind === "pageBreak" || block.kind === "columnBreak" || block.kind === "sectionBreak";
