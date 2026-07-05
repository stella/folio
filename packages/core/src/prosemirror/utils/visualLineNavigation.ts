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
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { findVerticalScrollParent } from "../../utils/findVerticalScrollParent";

const CONTENT_LINE_SELECTOR = ".layout-page-content .layout-line";

/** @internal */
export type VisualLineState = {
  stickyX: number | null;
  lastVisualLineIndex: number;
}

/** @internal */
export function createVisualLineState(): VisualLineState {
  return { stickyX: null, lastVisualLineIndex: -1 };
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
  const spans = container.querySelectorAll("span[data-pm-start][data-pm-end]");
  for (const span of Array.from(spans)) {
    const spanEl = span as HTMLElement;
    const pmStart = Number(spanEl.dataset["pmStart"]);
    const pmEnd = Number(spanEl.dataset["pmEnd"]);
    if (spanEl.classList.contains("layout-run-tab")) {
      if (pmPos >= pmStart && pmPos < pmEnd) return spanEl.getBoundingClientRect().left;
      continue;
    }
    if (pmPos >= pmStart && pmPos <= pmEnd && span.firstChild?.nodeType === Node.TEXT_NODE) {
      const textNode = span.firstChild as Text;
      const charIndex = Math.min(pmPos - pmStart, textNode.length);
      const ownerDoc = spanEl.ownerDocument;
      if (!ownerDoc) continue;
      const range = ownerDoc.createRange();
      range.setStart(textNode, charIndex);
      range.setEnd(textNode, charIndex);
      return range.getBoundingClientRect().left;
    }
  }
  const emptyRuns = container.querySelectorAll(".layout-empty-run");
  for (const emptyRun of Array.from(emptyRuns)) {
    const paragraph = emptyRun.closest(".layout-paragraph") as HTMLElement;
    if (!paragraph) continue;
    const pmStart = Number(paragraph.dataset["pmStart"]);
    const pmEnd = Number(paragraph.dataset["pmEnd"]);
    if (pmPos >= pmStart && pmPos <= pmEnd) return emptyRun.getBoundingClientRect().left;
  }
  return null;
}

/** @internal */
export function findLineElementAtPosition(
  container: HTMLElement,
  pmPos: number,
): HTMLElement | null {
  const allLines = container.querySelectorAll(CONTENT_LINE_SELECTOR);
  for (const line of Array.from(allLines)) {
    const lineEl = line as HTMLElement;
    const spans = lineEl.querySelectorAll("span[data-pm-start][data-pm-end]");
    for (const span of Array.from(spans)) {
      const s = span as HTMLElement;
      const start = Number(s.dataset["pmStart"]);
      const end = Number(s.dataset["pmEnd"]);
      if (pmPos >= start && pmPos <= end) return lineEl;
    }
  }
  for (const line of Array.from(allLines)) {
    const lineEl = line as HTMLElement;
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
      const textNode = spanEl.firstChild;
      if (!textNode || textNode.nodeType !== Node.TEXT_NODE) return pmStart;
      const text = textNode as Text;
      const ownerDoc = spanEl.ownerDocument;
      if (!ownerDoc) return pmStart;
      let lo = 0;
      let hi = text.length;
      while (lo < hi) {
        const mid = Math.floor((lo + hi) / 2);
        const r = ownerDoc.createRange();
        r.setStart(text, mid);
        r.setEnd(text, mid);
        if (clientX < r.getBoundingClientRect().left) hi = mid;
        else lo = mid + 1;
      }
      if (lo > 0 && lo <= text.length) {
        const r = ownerDoc.createRange();
        r.setStart(text, lo - 1);
        r.setEnd(text, lo - 1);
        const leftX = r.getBoundingClientRect().left;
        r.setStart(text, Math.min(lo, text.length));
        r.setEnd(text, Math.min(lo, text.length));
        const rightX = r.getBoundingClientRect().left;
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
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") {
    if (
      ["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key) ||
      (event.key.length === 1 && !event.ctrlKey && !event.metaKey)
    ) {
      state.stickyX = null;
      state.lastVisualLineIndex = -1;
    }
    return false;
  }
  if (event.ctrlKey || event.metaKey) {
    state.stickyX = null;
    state.lastVisualLineIndex = -1;
    return false;
  }
  if (!container) return false;

  const allLines = Array.from(container.querySelectorAll(CONTENT_LINE_SELECTOR));
  if (allLines.length === 0) return false;

  const { from, anchor } = view.state.selection;

  if (state.stickyX === null) {
    const clientX = getCaretClientX(container, from);
    if (clientX === null) return false;
    state.stickyX = clientX;
  }

  let currentIndex: number;
  if (state.lastVisualLineIndex >= 0 && state.lastVisualLineIndex < allLines.length) {
    currentIndex = state.lastVisualLineIndex;
  } else {
    const currentLine = findLineElementAtPosition(container, from);
    if (!currentLine) return false;
    currentIndex = allLines.indexOf(currentLine);
    if (currentIndex === -1) return false;
  }

  const targetIndex = event.key === "ArrowUp" ? currentIndex - 1 : currentIndex + 1;
  if (targetIndex < 0 || targetIndex >= allLines.length) {
    state.lastVisualLineIndex = -1;
    return false;
  }

  const targetLine = allLines[targetIndex] as HTMLElement;
  const newPos = findPositionOnLineAtClientX(targetLine, state.stickyX);
  if (newPos === null) return false;

  state.lastVisualLineIndex = targetIndex;

  const { state: pmState, dispatch } = view;
  const clampedPos = Math.max(0, Math.min(newPos, pmState.doc.content.size));

  try {
    const sel = event.shiftKey
      ? TextSelection.create(pmState.doc, anchor, clampedPos)
      : TextSelection.create(pmState.doc, clampedPos);
    dispatch(pmState.tr.setSelection(sel));
  } catch {
    const $newPos = pmState.doc.resolve(clampedPos);
    const sel = event.shiftKey
      ? TextSelection.between(pmState.doc.resolve(anchor), $newPos)
      : TextSelection.near($newPos);
    dispatch(pmState.tr.setSelection(sel));
  }

  scrollIntoViewIfNeeded(targetLine);
  return true;
}
