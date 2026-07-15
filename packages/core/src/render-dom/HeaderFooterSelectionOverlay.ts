import {
  findHfCaretSpan,
  findHfPmSpans,
  type HfSlotKind,
} from "../layout-bridge/dom/findHfPmSpans";

const CARET_CLASS = "folio-hf-selection-caret";
const RANGE_CLASS = "folio-hf-selection-rect";
const OWNED_SELECTOR = `.${CARET_CLASS}, .${RANGE_CLASS}`;
const TEXT_NODE_TYPE = 3;

export type HeaderFooterSelection = {
  from: number;
  kind: HfSlotKind;
  pageNumber?: number;
  rId: string;
  to: number;
};

export type HeaderFooterSelectionRect = {
  height: number;
  width: number;
  x: number;
  y: number;
};

export type HeaderFooterSelectionGeometry = {
  caret: Omit<HeaderFooterSelectionRect, "width"> | null;
  ranges: HeaderFooterSelectionRect[];
};

const relativeRect = (
  rect: Pick<DOMRect, "height" | "left" | "top" | "width">,
  containerRect: DOMRect,
  zoom: number,
): HeaderFooterSelectionRect => ({
  x: (rect.left - containerRect.left) / zoom,
  y: (rect.top - containerRect.top) / zoom,
  width: rect.width / zoom,
  height: rect.height / zoom,
});

/** Resolve a persistent header/footer PM selection against its painted DOM. */
export const resolveHeaderFooterSelectionGeometry = (
  pagesContainer: HTMLElement,
  selection: HeaderFooterSelection,
  zoom: number,
): HeaderFooterSelectionGeometry => {
  const containerRect = pagesContainer.getBoundingClientRect();
  const zoomDivisor = zoom === 0 ? 1 : zoom;
  const pageScope: ParentNode = selection.pageNumber
    ? (pagesContainer.querySelector(`.layout-page[data-page-number="${selection.pageNumber}"]`) ??
      pagesContainer)
    : pagesContainer;

  if (selection.from === selection.to) {
    const hit = findHfCaretSpan(pageScope, selection.kind, selection.rId, selection.from);
    if (!hit) {
      return { caret: null, ranges: [] };
    }
    const anchorRect = hit.element.getBoundingClientRect();
    let left = hit.edge === "right" ? anchorRect.right : anchorRect.left;
    let top = anchorRect.top;
    let height = anchorRect.height || 16;
    const pmStart = Number.parseInt(hit.element.dataset["pmStart"] ?? "", 10);
    const pmEnd = Number.parseInt(hit.element.dataset["pmEnd"] ?? "", 10);
    const textNode = hit.element.firstChild;
    if (
      textNode?.nodeType === TEXT_NODE_TYPE &&
      Number.isFinite(pmStart) &&
      Number.isFinite(pmEnd)
    ) {
      const offset = Math.min(
        Math.max(0, selection.from - pmStart),
        textNode.textContent?.length ?? 0,
      );
      const range = hit.element.ownerDocument.createRange();
      range.setStart(textNode, offset);
      range.setEnd(textNode, offset);
      const rangeRect = range.getBoundingClientRect();
      if (rangeRect.height > 0 || rangeRect.width > 0 || rangeRect.left > 0) {
        left = rangeRect.left;
        top = rangeRect.top;
        height = rangeRect.height || height;
      }
    }
    return {
      caret: {
        x: (left - containerRect.left) / zoomDivisor,
        y: (top - containerRect.top) / zoomDivisor,
        height: height / zoomDivisor,
      },
      ranges: [],
    };
  }

  const from = Math.min(selection.from, selection.to);
  const to = Math.max(selection.from, selection.to);
  const ranges: HeaderFooterSelectionRect[] = [];
  for (const span of findHfPmSpans(pageScope, selection.kind, selection.rId)) {
    const spanStart = Number.parseInt(span.dataset["pmStart"] ?? "", 10);
    const spanEnd = Number.parseInt(span.dataset["pmEnd"] ?? "", 10);
    if (!Number.isFinite(spanStart) || !Number.isFinite(spanEnd)) continue;
    if (spanEnd <= from || spanStart >= to) continue;

    const textNode = span.firstChild;
    if (textNode?.nodeType === TEXT_NODE_TYPE) {
      const length = textNode.textContent?.length ?? 0;
      const startOffset = Math.max(0, from - spanStart);
      const endOffset = Math.min(length, to - spanStart);
      if (startOffset < endOffset) {
        const range = span.ownerDocument.createRange();
        range.setStart(textNode, startOffset);
        range.setEnd(textNode, endOffset);
        for (const rect of Array.from(range.getClientRects())) {
          ranges.push(relativeRect(rect, containerRect, zoomDivisor));
        }
        continue;
      }
    }
    ranges.push(relativeRect(span.getBoundingClientRect(), containerRect, zoomDivisor));
  }
  return { caret: null, ranges };
};

const createOverlayElement = (pagesContainer: HTMLElement, className: string): HTMLDivElement => {
  const element = pagesContainer.ownerDocument.createElement("div");
  element.className = className;
  element.style.position = "absolute";
  element.style.pointerEvents = "none";
  return element;
};

export class HeaderFooterSelectionOverlay {
  clear(pagesContainer: HTMLElement): void {
    for (const element of pagesContainer.querySelectorAll(OWNED_SELECTOR)) {
      element.remove();
    }
  }

  sync(pagesContainer: HTMLElement, selection: HeaderFooterSelection | null, zoom: number): void {
    this.clear(pagesContainer);
    if (!selection) return;
    const geometry = resolveHeaderFooterSelectionGeometry(pagesContainer, selection, zoom);
    for (const rect of geometry.ranges) {
      const element = createOverlayElement(pagesContainer, RANGE_CLASS);
      element.dataset["testid"] = "hf-selection-rect";
      element.style.left = `${rect.x}px`;
      element.style.top = `${rect.y}px`;
      element.style.width = `${rect.width}px`;
      element.style.height = `${rect.height}px`;
      element.style.background = "var(--doc-selection, rgba(66, 133, 244, 0.3))";
      element.style.opacity = "0.35";
      element.style.zIndex = "10";
      pagesContainer.append(element);
    }
    if (!geometry.caret) return;
    const caret = createOverlayElement(pagesContainer, CARET_CLASS);
    caret.dataset["testid"] = "hf-caret";
    caret.style.left = `${geometry.caret.x}px`;
    caret.style.top = `${geometry.caret.y}px`;
    caret.style.width = "2px";
    caret.style.height = `${geometry.caret.height}px`;
    caret.style.background = "var(--doc-canvas-text, #000)";
    caret.style.zIndex = "11";
    caret.style.animation = "folio-caret-blink 1060ms steps(1, end) infinite";
    pagesContainer.append(caret);
  }
}
