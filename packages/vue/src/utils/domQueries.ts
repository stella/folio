/**
 * DOM-side hit-test and scroll helpers for the Vue editor — find the
 * painted PM span containing a position, scroll a position into view,
 * resolve a click coordinate back to a PM position, and the
 * double-/triple-click word/paragraph selection helpers.
 *
 * Every function takes containers as parameters; nothing closes over a
 * Vue ref. The selection helpers take a `setPmSelection` callback so the
 * caller controls how the resulting range gets dispatched to PM.
 *
 * Vue port of the upstream docx-editor util. `findBodyPmSpans` /
 * `findBodyPmAnchor` / `clickToPositionDom` resolve to our fork's split
 * `layout-bridge/dom/*` modules; `findWordBoundaries` is the Vue-package
 * text-selection util (core has no equivalent).
 */

import type { EditorView } from "prosemirror-view";
import {
  findBodyPmSpans,
  findBodyPmAnchor,
} from "@stll/folio-core/layout-bridge/dom/findBodyPmSpans";
import { clickToPositionDom } from "@stll/folio-core/layout-bridge/dom/clickToPositionDom";
import { findWordBoundaries } from "./textSelection";

/**
 * Resolve the painted header/footer instance nearest the viewport center — the
 * one the user is editing. The same HF is painted on every page (shared
 * `r:id`), so the chrome outline must track the active page rather than snap to
 * page one's copy (mirrors core's `getHfDomSnapshot` host pick; #691).
 */
export function nearestHfHostEl(position: "header" | "footer"): HTMLElement | null {
  const hosts = window.document.querySelectorAll<HTMLElement>(`.layout-page-${position}`);
  const first = hosts[0];
  if (!first) return null;
  const vpCenter = window.innerHeight / 2;
  let host = first;
  let bestDist = Infinity;
  for (const h of Array.from(hosts)) {
    const r = h.getBoundingClientRect();
    const dist = Math.abs((r.top + r.bottom) / 2 - vpCenter);
    if (dist < bestDist) {
      bestDist = dist;
      host = h;
    }
  }
  return host;
}

/**
 * Find the painted span containing `pmPos`. By default scoped to body spans
 * (which carry both pmStart and pmEnd) so HF runs in the separate PM document
 * don't mis-resolve double-/triple-click selection. When `hfSection` is set
 * (the user is editing a header/footer), scope to that section's painted host
 * instead — the HF spans live in `.layout-page-header` / `.layout-page-footer`
 * and carry PM positions from the HF document, so resolving against body spans
 * would compute word/paragraph bounds from the wrong text (#691).
 */
export function findElementAtPosition(
  container: HTMLElement,
  pmPos: number,
  hfSection?: "header" | "footer",
): HTMLElement | null {
  const els = hfSection
    ? // The same HF doc is painted on every page; the first host's spans share
      // the HF PM coord space, so one host suffices.
      Array.from(
        container.querySelectorAll<HTMLElement>(
          `.layout-page-${hfSection} span[data-pm-start][data-pm-end]`,
        ),
      )
    : findBodyPmSpans(container);
  for (const el of els) {
    const start = Number(el.dataset["pmStart"]);
    const end = Number(el.dataset["pmEnd"]);
    if (!Number.isNaN(start) && !Number.isNaN(end) && pmPos >= start && pmPos <= end) {
      return el;
    }
  }
  return null;
}

/**
 * Smooth-scroll the viewport so the painted element at `pmPos` is
 * visible (48px top padding). Falls back to a CSS attribute selector
 * when no body span carries pmPos in its [start,end] range.
 */
export function scrollVisiblePositionIntoView(
  pagesContainer: HTMLElement | null,
  viewport: HTMLElement | null,
  pmPos: number,
): void {
  if (!pagesContainer || !viewport) return;
  // Resolve the painted element the same way the React paged-scroll API does:
  // an exact, body-scoped `data-pm-start` anchor first (paragraph elements,
  // including headings, carry one), then the run-span [start,end] range. A
  // heading's pmPos is the paragraph node position, which only the anchor
  // match catches — the span loop alone would miss it (#930). The old
  // unscoped `[data-pm-start]` fallback is dropped: it could latch onto a
  // header/footer element sharing the same PM position.
  let targetEl: HTMLElement | null = findBodyPmAnchor(pagesContainer, pmPos);
  if (!targetEl) {
    for (const el of findBodyPmSpans(pagesContainer)) {
      const start = Number(el.dataset["pmStart"]);
      const end = Number(el.dataset["pmEnd"]);
      if (Number.isFinite(start) && Number.isFinite(end) && pmPos >= start && pmPos <= end) {
        targetEl = el;
        break;
      }
    }
  }
  if (!targetEl) return;
  const viewportRect = viewport.getBoundingClientRect();
  const targetRect = targetEl.getBoundingClientRect();
  viewport.scrollTo({
    top: targetRect.top - viewportRect.top + viewport.scrollTop - 48,
    behavior: "smooth",
  });
}

/**
 * Resolve a viewport-space click coordinate to a PM document position,
 * clamped to `doc.content.size`.
 */
export function resolvePos(
  pagesContainer: HTMLElement | null,
  view: EditorView | null,
  clientX: number,
  clientY: number,
): number | null {
  if (!pagesContainer || !view) return null;
  const pos = clickToPositionDom(pagesContainer, clientX, clientY, 1);
  if (pos === null || pos < 0) return null;
  return Math.min(pos, view.state.doc.content.size);
}

/**
 * Double-click word selection — expand `pos` to its word bounds and
 * hand the resulting range to `setPmSelection`.
 */
export function selectWord(
  pagesContainer: HTMLElement | null,
  pos: number,
  setPmSelection: (from: number, to: number) => void,
  hfSection?: "header" | "footer",
): void {
  if (!pagesContainer) return;
  const el = findElementAtPosition(pagesContainer, pos, hfSection);
  if (!el) return;
  const text = el.textContent ?? "";
  const pmStart = Number(el.dataset["pmStart"]) || 0;
  const offset = pos - pmStart;
  const [start, end] = findWordBoundaries(text, offset);
  const from = pmStart + start;
  const to = pmStart + end;
  if (from < to) {
    setPmSelection(from, to);
  }
}

/**
 * Triple-click paragraph selection — expand `pos` to the enclosing
 * `.layout-paragraph` element's PM range.
 */
export function selectParagraph(
  pagesContainer: HTMLElement | null,
  pos: number,
  setPmSelection: (from: number, to: number) => void,
  hfSection?: "header" | "footer",
): void {
  if (!pagesContainer) return;
  const el = findElementAtPosition(pagesContainer, pos, hfSection);
  if (!el) return;
  const paragraph = el.closest<HTMLElement>(".layout-paragraph");
  if (!paragraph) return;
  const pmStart = Number(paragraph.dataset["pmStart"]);
  const pmEnd = Number(paragraph.dataset["pmEnd"]);
  if (!Number.isNaN(pmStart) && !Number.isNaN(pmEnd) && pmStart < pmEnd) {
    setPmSelection(pmStart, pmEnd);
  }
}
