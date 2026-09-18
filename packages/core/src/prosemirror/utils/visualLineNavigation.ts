/**
 * Visual line navigation helpers — implements Word/Google-Docs-style
 * ArrowUp / ArrowDown with sticky X across visual lines (not just
 * paragraphs). Lifted from packages/react/src/paged-editor/
 * useVisualLineNavigation.ts so both adapters share the algorithm.
 *
 * Frontend-agnostic: takes a `getContainer: () => HTMLElement | null`
 * callback and a mutable sticky-state object, returns the same
 * function quartet React's hook returns.
 *
 * @packageDocumentation
 * @internal
 */
import { Selection, TextSelection } from "prosemirror-state";
import type { Node as PmNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { findCollapsedLineEdgeCaretTarget } from "../../layout-bridge/dom/clickToPositionDom";
import { findBodyEmptyRuns, findBodyPmSpans } from "../../layout-bridge/dom/findBodyPmSpans";
import { findVerticalScrollParent } from "../../utils/findVerticalScrollParent";
import {
  createTextStreamRange,
  descendantTextNodes,
  totalTextLength,
} from "../../layout-bridge/dom/textStreamDom";

const CONTENT_LINE_SELECTOR = ".layout-page-content .layout-line";
const PM_SPAN_SELECTOR = "span[data-pm-start][data-pm-end]";

/**
 * Where the previous vertical step left the caret: the column later steps keep
 * (`stickyX`), the visual line it settled on, and the caret position both
 * describe. The line index is not redundant with the position — a soft-wrap
 * boundary position belongs to two visual lines (the end of the wrapped line
 * and the start of the next), so the painted DOM alone cannot say which one the
 * caret sits on.
 *
 * @internal
 */
type VisualLineStep = {
  stickyX: number;
  lineIndex: number;
  pmPos: number;
};

/** @internal */
export type VisualLineState = {
  /**
   * Only describes the caret while it is still at `step.pmPos`. Any other move
   * (a click, a find result, an agent edit, typing) leaves the caret elsewhere,
   * and the next vertical step re-resolves both column and line from the DOM
   * instead of continuing from a line the caret has left.
   */
  step: VisualLineStep | null;
};

/** @internal */
export function createVisualLineState(): VisualLineState {
  return { step: null };
}

function scrollIntoViewIfNeeded(el: HTMLElement): void {
  const container = findVerticalScrollParent(el);
  if (!container) return;
  const elRect = el.getBoundingClientRect();
  const containerRect = container.getBoundingClientRect();
  const margin = 40;
  if (elRect.bottom > containerRect.bottom - margin) {
    container.scrollTop += elRect.bottom - containerRect.bottom + margin;
  } else if (elRect.top < containerRect.top + margin) {
    container.scrollTop -= containerRect.top - elRect.top + margin;
  }
}

/** @internal */
export function getCaretClientX(container: HTMLElement, pmPos: number): number | null {
  const spans = findBodyPmSpans(container);
  // A line-edge space is painted at zero font size, so a range inside it
  // reports the column *before* the space while the caret is painted after it.
  // Share the painted caret's geometry so a vertical step keeps the column the
  // user sees. (Same resolver order as `getCaretPositionFromDom`.)
  const collapsedTarget = findCollapsedLineEdgeCaretTarget(spans, pmPos);
  if (collapsedTarget) return collapsedTarget.geometry.left;
  for (const spanEl of spans) {
    const pmStart = Number(spanEl.dataset["pmStart"]);
    const pmEnd = Number(spanEl.dataset["pmEnd"]);
    if (spanEl.classList.contains("layout-run-tab")) {
      if (pmPos >= pmStart && pmPos < pmEnd) return spanEl.getBoundingClientRect().left;
      continue;
    }
    if (pmPos >= pmStart && pmPos <= pmEnd) {
      const textNodes = descendantTextNodes(spanEl);
      const charIndex = Math.min(pmPos - pmStart, totalTextLength(textNodes));
      const range = createTextStreamRange(spanEl, charIndex, charIndex);
      if (range) return range.getBoundingClientRect().left;
    }
  }
  for (const emptyRun of findBodyEmptyRuns(container)) {
    const paragraph = emptyRun.closest(".layout-paragraph") as HTMLElement;
    if (!paragraph) continue;
    const pmStart = Number(paragraph.dataset["pmStart"]);
    const pmEnd = Number(paragraph.dataset["pmEnd"]);
    if (pmPos >= pmStart && pmPos <= pmEnd) return emptyRun.getBoundingClientRect().left;
  }
  return null;
}

/**
 * Whether one painted line carries `pmPos`. Span endpoints are inclusive, so a
 * soft-wrap boundary answers true for both the line it ends and the line it
 * starts.
 */
const lineCarriesPosition = (lineEl: HTMLElement, pmPos: number): boolean => {
  for (const span of Array.from(lineEl.querySelectorAll<HTMLElement>(PM_SPAN_SELECTOR))) {
    const start = Number(span.dataset["pmStart"]);
    const end = Number(span.dataset["pmEnd"]);
    if (pmPos >= start && pmPos <= end) return true;
  }
  return false;
};

/** @internal */
export function findLineElementAtPosition(
  container: HTMLElement,
  pmPos: number,
): HTMLElement | null {
  const allLines = container.querySelectorAll<HTMLElement>(CONTENT_LINE_SELECTOR);
  for (const lineEl of Array.from(allLines)) {
    if (lineCarriesPosition(lineEl, pmPos)) return lineEl;
  }
  for (const lineEl of Array.from(allLines)) {
    const paragraph = lineEl.closest(".layout-paragraph") as HTMLElement;
    if (!paragraph) continue;
    const pStart = Number(paragraph.dataset["pmStart"]);
    const pEnd = Number(paragraph.dataset["pmEnd"]);
    if (pmPos >= pStart && pmPos <= pEnd) {
      const firstLineOfParagraph = paragraph.querySelector(".layout-line");
      if (firstLineOfParagraph === lineEl) return lineEl;
    }
  }
  return null;
}

/** @internal */
export function findPositionOnLineAtClientX(lineEl: HTMLElement, clientX: number): number | null {
  const spans = lineEl.querySelectorAll("span[data-pm-start][data-pm-end]");
  const emptyRun = lineEl.querySelector<HTMLElement>(".layout-empty-run");
  if (emptyRun) {
    const emptyRunStart = emptyRun.dataset["pmStart"];
    if (emptyRunStart !== undefined) return Number(emptyRunStart);

    const paragraph = emptyRun.closest<HTMLElement>(".layout-paragraph");
    const paragraphStart = paragraph?.dataset["pmStart"];
    return paragraphStart === undefined ? null : Number(paragraphStart) + 1;
  }
  if (spans.length === 0) {
    const paragraph = lineEl.closest(".layout-paragraph") as HTMLElement;
    if (paragraph?.dataset["pmStart"]) return Number(paragraph.dataset["pmStart"]) + 1;
    return null;
  }
  for (const span of Array.from(spans)) {
    const spanEl = span as HTMLElement;
    const rect = spanEl.getBoundingClientRect();
    const pmStart = Number(spanEl.dataset["pmStart"]);
    const pmEnd = Number(spanEl.dataset["pmEnd"]);
    if (spanEl.classList.contains("layout-run-tab")) {
      if (clientX >= rect.left && clientX <= rect.right) {
        const mid = (rect.left + rect.right) / 2;
        return clientX < mid ? pmStart : pmEnd;
      }
      continue;
    }
    if (clientX >= rect.left && clientX <= rect.right) {
      const textNodes = descendantTextNodes(spanEl);
      const textLength = totalTextLength(textNodes);
      if (textLength === 0) return pmStart;
      let lo = 0;
      let hi = textLength;
      while (lo < hi) {
        const mid = Math.floor((lo + hi) / 2);
        const r = createTextStreamRange(spanEl, mid, mid);
        if (!r) return pmStart;
        if (clientX < r.getBoundingClientRect().left) hi = mid;
        else lo = mid + 1;
      }
      if (lo > 0 && lo <= textLength) {
        const left = createTextStreamRange(spanEl, lo - 1, lo - 1);
        const right = createTextStreamRange(
          spanEl,
          Math.min(lo, textLength),
          Math.min(lo, textLength),
        );
        if (!left || !right) return pmStart;
        const leftX = left.getBoundingClientRect().left;
        const rightX = right.getBoundingClientRect().left;
        if (Math.abs(clientX - leftX) < Math.abs(clientX - rightX)) {
          return pmStart + (lo - 1);
        }
      }
      return pmStart + Math.min(lo, pmEnd - pmStart);
    }
  }
  let closestSpan: HTMLElement | null = null;
  let closestDist = Infinity;
  for (const span of Array.from(spans)) {
    const spanEl = span as HTMLElement;
    const rect = spanEl.getBoundingClientRect();
    const dist = clientX < rect.left ? rect.left - clientX : clientX - rect.right;
    if (dist < closestDist) {
      closestDist = dist;
      closestSpan = spanEl;
    }
  }
  if (!closestSpan) return null;
  const rect = closestSpan.getBoundingClientRect();
  return clientX < rect.left
    ? Number(closestSpan.dataset["pmStart"])
    : Number(closestSpan.dataset["pmEnd"]);
}

type ResolveStepOptions = {
  state: VisualLineState;
  container: HTMLElement;
  allLines: readonly HTMLElement[];
  pmPos: number;
};

/**
 * The step the next vertical move continues from: the remembered one while the
 * caret is still where that move left it, the caret's own painted geometry
 * otherwise.
 */
const resolveVisualLineStep = ({
  state,
  container,
  allLines,
  pmPos,
}: ResolveStepOptions): VisualLineStep | null => {
  const remembered = state.step;
  if (remembered?.pmPos === pmPos) {
    const rememberedLine = allLines[remembered.lineIndex];
    if (rememberedLine && lineCarriesPosition(rememberedLine, pmPos)) return remembered;
  }

  const currentLine = findLineElementAtPosition(container, pmPos);
  if (!currentLine) return null;
  const lineIndex = allLines.indexOf(currentLine);
  if (lineIndex === -1) return null;
  const stickyX = getCaretClientX(container, pmPos);
  return stickyX === null ? null : { stickyX, lineIndex, pmPos };
};

type VerticalSelectionOptions = {
  doc: PmNode;
  anchor: number;
  head: number;
  extend: boolean;
};

const createVerticalSelection = ({
  doc,
  anchor,
  head,
  extend,
}: VerticalSelectionOptions): Selection => {
  try {
    return extend ? TextSelection.create(doc, anchor, head) : TextSelection.create(doc, head);
  } catch {
    const $head = doc.resolve(head);
    return extend ? TextSelection.between(doc.resolve(anchor), $head) : Selection.near($head);
  }
};

/**
 * Handle PM ArrowUp / ArrowDown with visual-line awareness + sticky
 * X. Returns true if the event was handled and PM should not run
 * its default behaviour. Mutates `state` so consecutive presses
 * keep the same sticky X.
 *
 * @internal
 */
export function handleVisualLineKeyDown(
  state: VisualLineState,
  view: EditorView,
  event: KeyboardEvent,
  container: HTMLElement | null,
): boolean {
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return false;
  if (event.ctrlKey || event.metaKey) return false;
  if (!container) return false;

  const allLines = Array.from(container.querySelectorAll<HTMLElement>(CONTENT_LINE_SELECTOR));
  if (allLines.length === 0) return false;

  // The head is where the caret is painted: for a shift-extended selection
  // `from` is the fixed end, which would step down from the wrong line.
  const { head, anchor } = view.state.selection;
  const step = resolveVisualLineStep({ state, container, allLines, pmPos: head });
  if (!step) return false;

  const targetIndex = event.key === "ArrowUp" ? step.lineIndex - 1 : step.lineIndex + 1;
  const targetLine = allLines[targetIndex];
  if (!targetLine) {
    state.step = null;
    return false;
  }

  const newPos = findPositionOnLineAtClientX(targetLine, step.stickyX);
  if (newPos === null) return false;

  const { state: pmState, dispatch } = view;
  const clampedPos = Math.max(0, Math.min(newPos, pmState.doc.content.size));
  const selection = createVerticalSelection({
    doc: pmState.doc,
    anchor,
    head: clampedPos,
    extend: event.shiftKey,
  });
  dispatch(pmState.tr.setSelection(selection));
  state.step = { stickyX: step.stickyX, lineIndex: targetIndex, pmPos: selection.head };

  scrollIntoViewIfNeeded(targetLine);
  return true;
}
