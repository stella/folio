/**
 * Measure/paint line parity, shared by the specs that compare the width the
 * measurer decided for each painted line with the width the browser drew.
 */

import type { Page } from "@playwright/test";

/** Per-line comparison, in CSS pixels at zoom 1. */
export type LineParity = {
  index: number;
  measured: number;
  painted: number;
  delta: number;
  text: string;
  /** Whether the line carries joiner furniture, i.e. a repaired face change. */
  hasJoiner: boolean;
  /** Content runs on the line; the rounding budget scales with it. */
  runCount: number;
};

/**
 * Sub-pixel slack. Canvas advances and layout advances round independently, and
 * a line accumulates one rounding per run, so the budget scales with run count
 * rather than being a flat number that silently absorbs a real regression.
 *
 * Counted over CONTENT runs only. Counting every child would let each joiner
 * span, list marker and break marker buy another half-pixel of slack, so a
 * repaired line would be judged more loosely than the plain line beside it.
 */
export const PER_RUN_TOLERANCE_PX = 0.5;

/**
 * Read every painted line and compare the measurer's claim with the drawn
 * extent. Runs entirely in the page: both numbers must come from one layout
 * pass, or a resize between them would be read as a divergence.
 */
export function collectLineParity(page: Page): Promise<LineParity[]> {
  return page.evaluate(() => {
    const lines = [...document.querySelectorAll<HTMLElement>(".layout-line")];
    const results: {
      index: number;
      measured: number;
      painted: number;
      delta: number;
      text: string;
      hasJoiner: boolean;
      runCount: number;
    }[] = [];

    lines.forEach((lineEl, index) => {
      const claimed = lineEl.dataset["measuredWidth"];
      if (claimed === undefined) return;
      const measured = Number(claimed);
      if (!Number.isFinite(measured) || measured <= 0) return;

      // Compare like with like: `MeasuredLine.width` covers the line's CONTENT
      // runs. Page furniture painted into the same element — the list marker on
      // a first line, the zero-width break marker, joiner spans — is positioned
      // by other means and is not in that number, so a Range over the whole
      // element would read the marker's width as a measurement error. Content
      // is exactly what carries pm positions.
      const content = [...lineEl.querySelectorAll<HTMLElement>("[data-pm-start]")];
      if (content.length === 0) return;
      const text = content.map((el) => el.textContent ?? "").join("");
      if (text.trim().length === 0) return;

      const range = document.createRange();
      // SAFETY: length checked above.
      range.setStartBefore(content[0]!);

      // Trailing whitespace at a soft wrap hangs past the line, exactly as Word
      // renders it, and the measurer excludes it from the line width on purpose
      // (`trimTrailingSpacesAndTabs`). End the range at the last non-whitespace
      // character so both sides describe the same span of text.
      let ended = false;
      for (let i = content.length - 1; i >= 0 && !ended; i--) {
        // SAFETY: i is inside the array bounds.
        const span = content[i]!;
        const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
        const textNodes: Text[] = [];
        let node = walker.nextNode();
        while (node !== null) {
          textNodes.push(node as Text);
          node = walker.nextNode();
        }
        for (let j = textNodes.length - 1; j >= 0; j--) {
          // SAFETY: j is inside the array bounds.
          const textNode = textNodes[j]!;
          const trimmed = (textNode.data ?? "").replace(/[\s\u200B]+$/u, "");
          if (trimmed.length === 0) continue;
          range.setEnd(textNode, trimmed.length);
          ended = true;
          break;
        }
      }
      if (!ended) return;

      const painted = range.getBoundingClientRect().width;
      range.detach();
      if (painted <= 0) return;

      results.push({
        index,
        measured,
        painted,
        delta: painted - measured,
        text: text.slice(0, 60),
        hasJoiner: lineEl.querySelector("[data-docx-joiner]") !== null,
        runCount: content.length,
      });
    });
    return results;
  });
}
