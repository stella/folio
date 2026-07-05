/**
 * Header / Footer caret + selection projection utilities.
 *
 * Viewport-relative caret and selection rects for a persistent HF EditorView's
 * selection, resolved against the painter's `data-pm-start` / `data-pm-end`
 * spans inside `.layout-page-header` / `.layout-page-footer`. Lives in core so
 * the React + Vue adapters share a single implementation (they used to carry
 * byte-identical copies).
 *
 * @packageDocumentation
 * @public
 */

import type { EditorView } from "prosemirror-view";

// ============================================================================
// HF DOM snapshot cache — shared by the caret + selection-rect computations
// ============================================================================

type HfDomSnapshot = {
  host: HTMLElement;
  spans: HTMLElement[];
  ranged: HTMLElement[];
};

// Resolved HF DOM snapshot cached between calls, keyed by section. Invalidated
// by the painter's `painter:painted` event (`invalidateHfDomCache()` below) so
// the snapshot is always at most one paint stale. Without this, every
// HF caret + selection-rect computation re-walked every span on every
// page, which on multi-page docs is O(pages × spans) per scroll-rAF.
//
// Keyed by section because the header and footer are distinct PM docs painted
// in distinct hosts. A single shared slot let the first match in DOM order
// (always the header) shadow the footer, so an active footer's caret/selection
// resolved against the header's spans (#671).
const hfDomCache: { header: HfDomSnapshot | null; footer: HfDomSnapshot | null } = {
  header: null,
  footer: null,
};

/**
 * Drop the cached HF host + span lists. Hosts/painters call this after
 * a repaint (or HF mode toggle) so the next caret / selection compute
 * re-walks the DOM. Public so adapters can call it from their painter
 * commit signal.
 *
 * @public
 */
export function invalidateHfDomCache(): void {
  hfDomCache.header = null;
  hfDomCache.footer = null;
}

function getHfDomSnapshot(
  section: "header" | "footer",
  doc: globalThis.Document,
): HfDomSnapshot | null {
  // The same HF doc is painted on every page (shared by `r:id`), so any painted
  // instance carries the right PM coords. But the caret/selection overlay must
  // render on the instance the user is actually editing — pick the host nearest
  // the viewport center. Always taking the first (page 1) host drew the overlay
  // on page 1 even while editing a header/footer on a later page, so the user
  // saw no caret or highlight where they were typing (#691 footer).
  // Scoping to `.layout-page-${section}` keeps the header and footer from
  // shadowing each other (#671).
  const hosts = doc.querySelectorAll<HTMLElement>(`.layout-page-${section}`);
  if (hosts.length === 0) return null;
  const win = doc.defaultView;
  const vpCenter = win ? win.innerHeight / 2 : 0;
  let host = hosts[0];
  if (!host) return null;
  let bestDist = Infinity;
  for (const h of Array.from(hosts)) {
    const r = h.getBoundingClientRect();
    const dist = Math.abs((r.top + r.bottom) / 2 - vpCenter);
    if (dist < bestDist) {
      bestDist = dist;
      host = h;
    }
  }
  // Reuse the cached span lists only when they belong to the same painted host
  // (and it's still live). The host changes as the user scrolls between pages,
  // so a section-only cache would keep resolving against the wrong instance.
  const cached = hfDomCache[section];
  if (cached && cached.host === host && cached.host.isConnected) return cached;
  const spans = Array.from(host.querySelectorAll<HTMLElement>("span[data-pm-start][data-pm-end]"));
  const ranged = Array.from(host.querySelectorAll<HTMLElement>("[data-pm-start][data-pm-end]"));
  const snapshot = { host, spans, ranged };
  hfDomCache[section] = snapshot;
  return snapshot;
}

// ============================================================================
// HF caret rect — used by both React and Vue adapters
// ============================================================================

/**
 * Viewport-relative caret rect for a persistent HF EditorView's selection
 * head. Resolves against the painter's `data-pm-start`/`data-pm-end` spans
 * inside `.layout-page-header` / `.layout-page-footer`. The same HF doc is
 * painted on every page (multi-page docs, titlePg), so this walks every
 * candidate host and picks the one whose spans bracket the PM head; falls
 * back to the first so empty paragraphs still resolve to a paragraph anchor.
 *
 * Public so the React + Vue adapters can share a single implementation
 * (`packages/{react,vue}` adapters used to carry byte-identical copies).
 *
 * @public
 */
export function computeHfCaretRectFromView(
  view: EditorView,
  section: "header" | "footer",
  doc: globalThis.Document = globalThis.document,
): { top: number; left: number; height: number } | null {
  const sel = view.state.selection;
  if (!sel.empty) return null;
  const pmPos = sel.head;
  const snapshot = getHfDomSnapshot(section, doc);
  if (!snapshot) return null;
  const { host, spans } = snapshot;
  for (const span of spans) {
    const start = Number(span.dataset["pmStart"]);
    const end = Number(span.dataset["pmEnd"]);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    if (pmPos >= start && pmPos <= end) {
      const range = host.ownerDocument.createRange();
      const walker = host.ownerDocument.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      let remaining = pmPos - start;
      let textNode = walker.nextNode() as Text | null;
      while (textNode) {
        const len = textNode.data.length;
        if (remaining <= len) {
          try {
            range.setStart(textNode, remaining);
            range.setEnd(textNode, remaining);
            const rect = range.getClientRects()[0] ?? range.getBoundingClientRect();
            if (rect && rect.height > 0) {
              return { top: rect.top, left: rect.left, height: rect.height };
            }
          } catch {
            // fall through
          }
          break;
        }
        remaining -= len;
        textNode = walker.nextNode() as Text | null;
      }
      const spanRect = span.getBoundingClientRect();
      const ratio = (pmPos - start) / Math.max(1, end - start);
      return {
        top: spanRect.top,
        left: spanRect.left + spanRect.width * ratio,
        height: spanRect.height,
      };
    }
  }
  // Exact paragraph/line anchor at `pmPos` (when the painter emits one).
  const anchor = host.querySelector<HTMLElement>(`[data-pm-start="${pmPos}"]`);
  if (anchor) {
    const rect = anchor.getBoundingClientRect();
    return { top: rect.top, left: rect.left + 1, height: rect.height || 16 };
  }

  // Fallback for empty paragraphs / line-ends: walk every painted element
  // that carries `[data-pm-start][data-pm-end]` and find the one whose
  // range brackets `pmPos`. Use its rect — left edge for an empty
  // paragraph (cursor at the paragraph's start), right edge if the cursor
  // is at the paragraph's end. Without this, hitting Enter into a new
  // empty paragraph hid the caret entirely until the user typed.
  const ranged = snapshot.ranged;
  let bestEl: HTMLElement | null = null;
  let bestSpan = Infinity;
  for (const el of ranged) {
    const start = Number(el.dataset["pmStart"]);
    const end = Number(el.dataset["pmEnd"]);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    if (pmPos < start || pmPos > end) continue;
    const span = end - start;
    if (span < bestSpan) {
      bestSpan = span;
      bestEl = el;
    }
  }
  if (bestEl) {
    const rect = bestEl.getBoundingClientRect();
    const end = Number(bestEl.dataset["pmEnd"]);
    const atEnd = pmPos >= end;
    return {
      top: rect.top,
      left: atEnd ? rect.right : rect.left + 1,
      height: rect.height || 16,
    };
  }

  // Cursor sits past every painted element's `[pmStart, pmEnd]` range —
  // typically because the cursor is at `doc.content.size` (end of last
  // paragraph). Find the painted element with the largest `pmStart` that
  // is still `<= pmPos` and snap the caret to its trailing edge. This is
  // a much better visual than "top-left of host" when the user has just
  // hit Enter to add a paragraph and is now sitting at the end of the
  // content.
  let trailingEl: HTMLElement | null = null;
  let trailingStart = -Infinity;
  for (const el of ranged) {
    const start = Number(el.dataset["pmStart"]);
    if (!Number.isFinite(start)) continue;
    if (start > pmPos) continue;
    if (start > trailingStart) {
      trailingStart = start;
      trailingEl = el;
    }
  }
  if (trailingEl) {
    const rect = trailingEl.getBoundingClientRect();
    return { top: rect.top, left: rect.right, height: rect.height || 16 };
  }

  // Last resort: anchor at the host's top-left so the caret is at least
  // visible while in HF edit mode. Better than disappearing.
  const hostRect = host.getBoundingClientRect();
  return {
    top: hostRect.top + 2,
    left: hostRect.left + 2,
    height: 16,
  };
}

/**
 * Selection-rect set for a non-empty HF selection, projected against the
 * painted HF spans. Mirror of `computeSelectionRectsFromDom` but scoped to
 * `.layout-page-header` / `.layout-page-footer` instead of the body. Used
 * so the painter draws a visible highlight when the user drag-selects
 * inside a header/footer in edit mode.
 *
 * Returns viewport-relative `{top, left, width, height}` rects. Empty
 * array when selection is collapsed or no painted spans overlap the range.
 *
 * @public
 */
export function computeHfSelectionRectsFromView(
  view: EditorView,
  section: "header" | "footer",
  doc: globalThis.Document = globalThis.document,
): Array<{ top: number; left: number; width: number; height: number }> {
  const sel = view.state.selection;
  if (sel.empty) return [];
  const from = sel.from;
  const to = sel.to;
  const out: Array<{ top: number; left: number; width: number; height: number }> = [];

  // Reuse the cached HF DOM snapshot for this section. Every painted HF host
  // for the section shares the same PM coord space (only one HF doc, painted N
  // times for the N pages), so a single host's spans suffice for selection
  // rects.
  const snapshot = getHfDomSnapshot(section, doc);
  if (!snapshot) return out;
  const { host, spans } = snapshot;
  for (const spanEl of spans) {
    const pmStart = Number(spanEl.dataset["pmStart"]);
    const pmEnd = Number(spanEl.dataset["pmEnd"]);
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
    if (pmEnd <= from || pmStart >= to) continue;

    // Tab spans: full-span highlight.
    if (spanEl.classList.contains("layout-run-tab")) {
      const rect = spanEl.getBoundingClientRect();
      out.push({ top: rect.top, left: rect.left, width: rect.width, height: rect.height });
      continue;
    }

    let textNode: Text | null = null;
    if (spanEl.firstChild?.nodeType === Node.TEXT_NODE) {
      textNode = spanEl.firstChild as Text;
    } else if (
      spanEl.firstChild?.nodeType === Node.ELEMENT_NODE &&
      (spanEl.firstChild as HTMLElement).tagName === "A" &&
      spanEl.firstChild.firstChild?.nodeType === Node.TEXT_NODE
    ) {
      textNode = spanEl.firstChild.firstChild as Text;
    }
    if (!textNode) continue;

    const startChar = Math.max(0, from - pmStart);
    const endChar = Math.min(textNode.length, to - pmStart);
    if (startChar >= endChar) continue;

    const range = host.ownerDocument.createRange();
    range.setStart(textNode, startChar);
    range.setEnd(textNode, endChar);
    for (const rect of Array.from(range.getClientRects())) {
      out.push({ top: rect.top, left: rect.left, width: rect.width, height: rect.height });
    }
  }

  return out;
}
