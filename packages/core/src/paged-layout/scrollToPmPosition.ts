/**
 * Scroll the visible paged layout to a ProseMirror position.
 *
 * The paged editor's editing PM view is hidden off-screen
 * (HiddenProseMirror); only the painted pages are visible. Scrolling
 * therefore cannot use `view.coordsAtPos` — those coordinates belong
 * to the hidden, unpaginated mirror, not the pages the user sees.
 * This module owns the painted-DOM scroll path: per-run anchors when
 * the target is in the rendered buffer, page shells (always present
 * under virtualization) when it is not.
 *
 * `PagedEditor` delegates its `scrollToPosition` ref method here;
 * external hosts that only hold the hidden `EditorView` (e.g. the AI
 * suggestion thread) use {@link scrollFolioPositionIntoView}.
 */

import type { EditorView } from "prosemirror-view";

import {
  findBodyPmAnchor,
  findBodyPmAnchors,
} from "../layout-bridge/dom/findBodyPmSpans";
import {
  findPageShellForPmPos,
  renderPagesAround,
} from "../layout-painter/renderPage";
import {
  getEditorScrollRoot,
  scrollEditorElementIntoView,
} from "./editorScrollRoot";
import { isValidPmScrollPosition } from "./scrollNavigation";

/** Class on the pages container `PagedEditor` paints into. */
export const PAGES_CONTAINER_CLASS = "paged-editor__pages";

/** Scroll visible pages to a ProseMirror position. */
export const scrollPagesToPmPosition = (
  pageContainer: HTMLElement,
  pmPos: number,
): void => {
  if (!isValidPmScrollPosition(pmPos)) {
    return;
  }

  // A per-run query on a virtualized doc only sees the rendered buffer, so a
  // target outside it is located through its page shell (always present, with
  // correct dimensions). That page and its neighbours are painted before the
  // scroll, so the viewport lands on content in one move and the pages in
  // between are never painted; the distance-aware default behavior makes a
  // far jump instant.
  const rendered =
    findBodyPmAnchor(pageContainer, pmPos) ??
    findAnchorContaining(pageContainer, pmPos);
  if (rendered) {
    scrollEditorElementIntoView(rendered);
    return;
  }

  const shellHit = findPageShellForPmPos(pageContainer, pmPos);
  if (!shellHit) {
    return;
  }
  const { element: shell } = shellHit;
  renderPagesAround(pageContainer, shell);
  scrollEditorElementIntoView(
    findBodyPmAnchor(shell, pmPos) ??
      findAnchorContaining(shell, pmPos) ??
      findAnchorBefore(shell, pmPos) ??
      shell,
  );
};

type AnchorRange = { element: HTMLElement; start: number; end: number };

const anchorRanges = (root: ParentNode): AnchorRange[] => {
  const ranges: AnchorRange[] = [];
  for (const element of findBodyPmAnchors(root)) {
    const start = Number(element.dataset["pmStart"]);
    if (Number.isNaN(start)) {
      continue;
    }
    const endAttr = element.dataset["pmEnd"];
    ranges.push({
      element,
      start,
      end: endAttr === undefined ? start : Number(endAttr),
    });
  }
  return ranges;
};

/** First painted run whose range holds `pmPos` (block positions rarely match a run start). */
const findAnchorContaining = (
  root: ParentNode,
  pmPos: number,
): HTMLElement | null => {
  for (const { element, start, end } of anchorRanges(root)) {
    if (start <= pmPos && pmPos <= end) {
      return element;
    }
  }
  return null;
};

/** The painted run starting closest before `pmPos`. */
const findAnchorBefore = (
  root: ParentNode,
  pmPos: number,
): HTMLElement | null => {
  let best: HTMLElement | null = null;
  let bestStart = Number.NEGATIVE_INFINITY;
  for (const { element, start } of anchorRanges(root)) {
    if (start <= pmPos && start > bestStart) {
      bestStart = start;
      best = element;
    }
  }
  return best;
};

/**
 * Scroll the paged layout that mirrors `view` so `pmPos` is centered
 * in the viewport. The hidden editing view sits inside the editor's
 * scroll container, so the painted pages are reachable from its DOM.
 * Returns false when no paged layout is found (non-paged host) so
 * callers can fall back to coordinate-based scrolling.
 */
export const scrollFolioPositionIntoView = (
  view: EditorView,
  pmPos: number,
): boolean => {
  const scrollContainer = getEditorScrollRoot(view.dom);
  const pageContainer =
    scrollContainer?.querySelector<HTMLElement>(`.${PAGES_CONTAINER_CLASS}`) ??
    null;
  if (!pageContainer) {
    return false;
  }
  scrollPagesToPmPosition(pageContainer, pmPos);
  return true;
};
