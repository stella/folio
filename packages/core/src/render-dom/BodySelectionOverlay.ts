/**
 * Imperative body-selection painter shared by framework adapters. The editable
 * ProseMirror view is hidden, so its caret, range selection, and table-cell
 * selection must be projected onto the paginated DOM.
 */

import { NodeSelection, type EditorState } from "prosemirror-state";

import { applyCellSelectionHighlight } from "../layout-bridge/cellSelectionHighlight";
import {
  getCaretPositionFromDom,
  getSelectionRectsFromDom,
} from "../layout-bridge/dom/clickToPositionDom";
import { findBodyPmAnchor } from "../layout-bridge/dom/findBodyPmSpans";
import { findImageElement } from "../layout-painter/imageLayout";

const CARET_CLASS = "folio-body-selection-caret";
const RANGE_CLASS = "folio-body-selection-rect";
const OWNED_OVERLAY_SELECTOR = `.${CARET_CLASS}, .${RANGE_CLASS}`;

export type BodySelectionOverlayResult =
  | { type: "text" }
  | { type: "image"; element: HTMLElement; pmPos: number };

export type SyncBodySelectionOverlayOptions = {
  pagesContainer: HTMLElement;
  state: EditorState;
  zoom: number;
  zIndex?: number;
  caretColor?: string;
  selectionColor?: string;
};

const setOverlayBox = (
  element: HTMLElement,
  rect: { x: number; y: number; width: number; height: number },
  pagesContainer: HTMLElement,
  zoom: number,
): void => {
  const divisor = zoom > 0 ? zoom : 1;
  element.style.left = `${rect.x / divisor + pagesContainer.scrollLeft}px`;
  element.style.top = `${rect.y / divisor + pagesContainer.scrollTop}px`;
  element.style.width = `${rect.width / divisor}px`;
  element.style.height = `${rect.height / divisor}px`;
};

export class BodySelectionOverlay {
  clear(pagesContainer: HTMLElement): void {
    for (const element of pagesContainer.querySelectorAll(OWNED_OVERLAY_SELECTOR)) {
      element.remove();
    }
  }

  sync({
    pagesContainer,
    state,
    zoom,
    zIndex = 10,
    caretColor = "var(--doc-caret, #000)",
    selectionColor = "var(--doc-selection, rgba(66, 133, 244, 0.3))",
  }: SyncBodySelectionOverlayOptions): BodySelectionOverlayResult {
    this.clear(pagesContainer);
    applyCellSelectionHighlight(pagesContainer, state);

    const { selection } = state;
    if (selection instanceof NodeSelection && selection.node.type.name === "image") {
      const anchor = findBodyPmAnchor(pagesContainer, selection.from);
      const element = anchor ? findImageElement(anchor) : null;
      if (element) {
        return { type: "image", element, pmPos: selection.from };
      }
    }

    if (selection.empty) {
      const caret = getCaretPositionFromDom(
        pagesContainer,
        selection.from,
        pagesContainer.getBoundingClientRect(),
      );
      if (!caret) {
        return { type: "text" };
      }
      const element = pagesContainer.ownerDocument.createElement("div");
      element.className = CARET_CLASS;
      element.dataset["folioCaretRect"] = "";
      element.style.position = "absolute";
      element.style.background = caretColor;
      element.style.pointerEvents = "none";
      element.style.zIndex = String(zIndex);
      element.style.animation = "folio-caret-blink 1060ms steps(1, end) infinite";
      setOverlayBox(
        element,
        { x: caret.x, y: caret.y, width: 2 * (zoom > 0 ? zoom : 1), height: caret.height },
        pagesContainer,
        zoom,
      );
      pagesContainer.appendChild(element);
      return { type: "text" };
    }

    const rects = getSelectionRectsFromDom(
      pagesContainer,
      selection.from,
      selection.to,
      pagesContainer.getBoundingClientRect(),
    );
    for (const rect of rects) {
      const element = pagesContainer.ownerDocument.createElement("div");
      element.className = RANGE_CLASS;
      element.dataset["folioSelectionRect"] = "";
      element.dataset["pageIndex"] = String(rect.pageIndex);
      element.style.position = "absolute";
      element.style.background = selectionColor;
      element.style.pointerEvents = "none";
      element.style.zIndex = String(zIndex);
      setOverlayBox(element, rect, pagesContainer, zoom);
      pagesContainer.appendChild(element);
    }
    return { type: "text" };
  }
}
