/**
 * DOM helpers for transient paragraph flashes on the painted layout surface.
 */

import type { ParagraphHighlightOptions } from "./paragraphFlashTypes";

export type { ParagraphHighlightOptions, ScrollToParaIdOptions } from "./paragraphFlashTypes";

/** Default color used by paragraph flashes. */
export const DEFAULT_PARAGRAPH_FLASH_COLOR = "rgba(255, 235, 59, 0.55)";
/** Default duration for paragraph flashes. */
export const DEFAULT_PARAGRAPH_FLASH_DURATION_MS = 1200;
/** CSS class applied to paragraph fragments during a transient flash. */
export const PARAGRAPH_FLASH_CLASS_NAME = "folio-paragraph-flash";

const timers = new WeakMap<HTMLElement, ReturnType<typeof setTimeout>>();

const escapeAttributeValue = (value: string): string => {
  const cssGlobal = globalThis as typeof globalThis & {
    CSS?: { escape?: (value: string) => string };
  };
  if (typeof cssGlobal.CSS?.escape === "function") {
    return cssGlobal.CSS.escape(value);
  }
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
};

const normalizedColor = (options?: ParagraphHighlightOptions): string => {
  const color = options?.color?.trim();
  return color || DEFAULT_PARAGRAPH_FLASH_COLOR;
};

const normalizedDurationMs = (options?: ParagraphHighlightOptions): number => {
  const duration = options?.durationMs;
  if (duration == null) {
    return DEFAULT_PARAGRAPH_FLASH_DURATION_MS;
  }
  if (!Number.isFinite(duration) || duration < 0) {
    return DEFAULT_PARAGRAPH_FLASH_DURATION_MS;
  }
  return duration;
};

/** Find all painted paragraph fragments with a stable `data-para-id`. */
export const findParagraphFragmentsByParaId = (root: ParentNode, paraId: string): HTMLElement[] => {
  if (!paraId || !paraId.trim()) {
    return [];
  }
  const escaped = escapeAttributeValue(paraId);
  return Array.from(
    root.querySelectorAll<HTMLElement>(`.layout-paragraph[data-para-id="${escaped}"]`),
  );
};

/** Apply a transient flash to a collection of paragraph elements. */
export const flashParagraphElements = (
  elements: Iterable<HTMLElement>,
  options?: ParagraphHighlightOptions,
): number => {
  let count = 0;
  const color = normalizedColor(options);
  const durationMs = normalizedDurationMs(options);

  for (const el of elements) {
    count++;
    const existingTimer = timers.get(el);
    if (existingTimer !== undefined) {
      clearTimeout(existingTimer);
    }

    el.classList.remove(PARAGRAPH_FLASH_CLASS_NAME);
    void el.offsetWidth;
    el.style.setProperty("--folio-paragraph-flash-color", color);
    el.style.setProperty("--folio-paragraph-flash-duration", `${durationMs}ms`);
    el.classList.add(PARAGRAPH_FLASH_CLASS_NAME);

    const timer = setTimeout(() => {
      el.classList.remove(PARAGRAPH_FLASH_CLASS_NAME);
      el.style.removeProperty("--folio-paragraph-flash-color");
      el.style.removeProperty("--folio-paragraph-flash-duration");
      timers.delete(el);
    }, durationMs);
    timers.set(el, timer);
  }

  return count;
};

/** Find paragraph fragments by `paraId` and flash them. */
export const flashParagraphFragmentsByParaId = (
  root: ParentNode,
  paraId: string,
  options?: ParagraphHighlightOptions,
): boolean => flashParagraphElements(findParagraphFragmentsByParaId(root, paraId), options) > 0;
