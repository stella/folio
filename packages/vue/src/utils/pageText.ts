import type { Node as PMNode } from "prosemirror-model";

import type { Layout } from "@stll/folio-core/layout-engine";

/**
 * The plain text of a single rendered page (1-based), read by joining each of
 * the page's layout fragments' `[pmStart, pmEnd)` document text with a
 * newline. Fragments with no `pmStart`/`pmEnd` (e.g. a page-shell placeholder)
 * are skipped. Returns `null` when the layout hasn't been computed yet or the
 * page number is out of range.
 *
 * Mirrors `@stll/folio-react`'s `components/pageText.ts`; kept as a separate
 * file rather than shared so `@stll/folio-vue` doesn't cross-import the React
 * package.
 */
export const getPageTextFromLayout = (
  layout: Layout | null,
  doc: PMNode,
  page: number,
): string | null => {
  if (!layout || !Number.isInteger(page) || page < 1) {
    return null;
  }
  const layoutPage = layout.pages[page - 1];
  if (!layoutPage) {
    return null;
  }
  const fragmentTexts: string[] = [];
  for (const fragment of layoutPage.fragments) {
    if (typeof fragment.pmStart !== "number" || typeof fragment.pmEnd !== "number") {
      continue;
    }
    fragmentTexts.push(doc.textBetween(fragment.pmStart, fragment.pmEnd, "\n"));
  }
  return fragmentTexts.join("\n");
};
