/**
 * Framework-neutral mapping from ProseMirror positions to the DOM emitted by
 * the layout painter. Framework adapters use this for decorations, anchored
 * chrome, and plugin surfaces without duplicating range geometry.
 */

import { findBodyEmptyRuns, findBodyPmSpans } from "../layout-bridge/dom/findBodyPmSpans";
import { closestHtmlElement } from "../utils/domGuards";

export type RenderedDomPoint = {
  x: number;
  y: number;
  height: number;
};

export type RenderedDomRect = {
  x: number;
  y: number;
  width: number;
  height: number;
};

export type RenderedDomContext = {
  getCoordinatesForPosition(pmPos: number): RenderedDomPoint | null;
  findElementsForRange(from: number, to: number): Element[];
  getRectsForRange(from: number, to: number): RenderedDomRect[];
  getContainerOffset(): { x: number; y: number };
};

const isTextNode = (node: ChildNode | null): node is Text => node?.nodeType === Node.TEXT_NODE;

const textNodeForSpan = (span: HTMLElement): Text | null => {
  const firstChild = span.firstChild;
  if (isTextNode(firstChild)) {
    return firstChild;
  }
  if (
    firstChild instanceof HTMLElement &&
    firstChild.tagName === "A" &&
    isTextNode(firstChild.firstChild)
  ) {
    return firstChild.firstChild;
  }
  return null;
};

const lineHeightFor = (element: Element, zoom: number): number => {
  const line = closestHtmlElement(element, ".layout-line");
  return line ? line.getBoundingClientRect().height / zoom : 16;
};

export class RenderedDomContextImpl implements RenderedDomContext {
  readonly #pagesContainer: HTMLElement;
  readonly #zoom: number;

  constructor(pagesContainer: HTMLElement, zoom = 1) {
    this.#pagesContainer = pagesContainer;
    this.#zoom = zoom > 0 ? zoom : 1;
  }

  getCoordinatesForPosition(pmPos: number): RenderedDomPoint | null {
    const containerRect = this.#pagesContainer.getBoundingClientRect();
    for (const span of findBodyPmSpans(this.#pagesContainer)) {
      const pmStart = Number(span.dataset["pmStart"]);
      const pmEnd = Number(span.dataset["pmEnd"]);
      const containsPosition = span.classList.contains("layout-run-tab")
        ? pmPos >= pmStart && pmPos < pmEnd
        : pmPos >= pmStart && pmPos <= pmEnd;
      if (!containsPosition) {
        continue;
      }

      const spanRect = span.getBoundingClientRect();
      if (span.classList.contains("layout-run-tab")) {
        return {
          x: (spanRect.left - containerRect.left) / this.#zoom,
          y: (spanRect.top - containerRect.top) / this.#zoom,
          height: lineHeightFor(span, this.#zoom),
        };
      }

      const textNode = textNodeForSpan(span);
      if (!textNode) {
        return {
          x: (spanRect.left - containerRect.left) / this.#zoom,
          y: (spanRect.top - containerRect.top) / this.#zoom,
          height: lineHeightFor(span, this.#zoom),
        };
      }

      const charIndex = Math.min(Math.max(0, pmPos - pmStart), textNode.length);
      const range = span.ownerDocument.createRange();
      range.setStart(textNode, charIndex);
      range.setEnd(textNode, charIndex);
      const rangeRect = range.getBoundingClientRect();
      return {
        x: (rangeRect.left - containerRect.left) / this.#zoom,
        y: (rangeRect.top - containerRect.top) / this.#zoom,
        height: lineHeightFor(span, this.#zoom),
      };
    }

    for (const emptyRun of findBodyEmptyRuns(this.#pagesContainer)) {
      const paragraph = closestHtmlElement(emptyRun, ".layout-paragraph");
      if (!paragraph) {
        continue;
      }
      const pmStart = Number(paragraph.dataset["pmStart"]);
      const pmEnd = Number(paragraph.dataset["pmEnd"]);
      if (pmPos < pmStart || pmPos > pmEnd) {
        continue;
      }
      const rect = emptyRun.getBoundingClientRect();
      return {
        x: (rect.left - containerRect.left) / this.#zoom,
        y: (rect.top - containerRect.top) / this.#zoom,
        height: lineHeightFor(emptyRun, this.#zoom),
      };
    }
    return null;
  }

  findElementsForRange(from: number, to: number): Element[] {
    return findBodyPmSpans(this.#pagesContainer).filter((span) => {
      const pmStart = Number(span.dataset["pmStart"]);
      const pmEnd = Number(span.dataset["pmEnd"]);
      return pmEnd > from && pmStart < to;
    });
  }

  getRectsForRange(from: number, to: number): RenderedDomRect[] {
    const containerRect = this.#pagesContainer.getBoundingClientRect();
    const rects: RenderedDomRect[] = [];
    for (const element of this.findElementsForRange(from, to)) {
      if (!(element instanceof HTMLElement)) {
        continue;
      }
      const pmStart = Number(element.dataset["pmStart"]);
      if (element.classList.contains("layout-run-tab")) {
        const rect = element.getBoundingClientRect();
        rects.push({
          x: (rect.left - containerRect.left) / this.#zoom,
          y: (rect.top - containerRect.top) / this.#zoom,
          width: rect.width / this.#zoom,
          height: rect.height / this.#zoom,
        });
        continue;
      }

      const textNode = textNodeForSpan(element);
      if (!textNode) {
        continue;
      }
      const startChar = Math.max(0, from - pmStart);
      const endChar = Math.min(textNode.length, to - pmStart);
      if (startChar >= endChar) {
        continue;
      }
      const range = element.ownerDocument.createRange();
      range.setStart(textNode, startChar);
      range.setEnd(textNode, endChar);
      for (const rect of Array.from(range.getClientRects())) {
        rects.push({
          x: (rect.left - containerRect.left) / this.#zoom,
          y: (rect.top - containerRect.top) / this.#zoom,
          width: rect.width / this.#zoom,
          height: rect.height / this.#zoom,
        });
      }
    }
    return rects;
  }

  getContainerOffset(): { x: number; y: number } {
    const parent = this.#pagesContainer.parentElement;
    if (!parent) {
      return { x: 0, y: 0 };
    }
    const containerRect = this.#pagesContainer.getBoundingClientRect();
    const parentRect = parent.getBoundingClientRect();
    return {
      x: (containerRect.left - parentRect.left) / this.#zoom,
      y: (containerRect.top - parentRect.top) / this.#zoom,
    };
  }
}

export const createRenderedDomContext = (
  pagesContainer: HTMLElement,
  zoom = 1,
): RenderedDomContext => new RenderedDomContextImpl(pagesContainer, zoom);
