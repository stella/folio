import type { Node as PMNode } from "prosemirror-model";

import type { MarkupViewProjection } from "../prosemirror/markupViewProjection";
import type { Layout } from "../layout-engine";

type PageTextOptions = {
  layout: Layout | null;
  doc: PMNode;
  page: number;
  projection?: MarkupViewProjection;
};

/**
 * The plain text of a single rendered page (1-based), read by joining each of
 * the page's layout fragments' `[pmStart, pmEnd)` document text with a
 * newline. A review projection maps those authored ranges to the displayed
 * document, excluding text the view hides. Fragments with no `pmStart`/`pmEnd`
 * (e.g. a page-shell placeholder)
 * are skipped. Returns `null` when the layout hasn't been computed yet or the
 * page number is out of range.
 */
export const getPageTextFromLayout = ({
  layout,
  doc,
  page,
  projection,
}: PageTextOptions): string | null => {
  if (!layout || !Number.isInteger(page) || page < 1) {
    return null;
  }
  const layoutPage = layout.pages[page - 1];
  if (!layoutPage) {
    return null;
  }
  const viewDoc = projection?.type === "resolved" ? projection.doc : doc;
  const fragmentTexts: string[] = [];
  for (const fragment of layoutPage.fragments) {
    if (typeof fragment.pmStart !== "number" || typeof fragment.pmEnd !== "number") {
      continue;
    }
    const from =
      projection?.type === "resolved" ? projection.toView(fragment.pmStart, 1) : fragment.pmStart;
    const to =
      projection?.type === "resolved" ? projection.toView(fragment.pmEnd, -1) : fragment.pmEnd;
    fragmentTexts.push(viewDoc.textBetween(from, to, "\n"));
  }
  return fragmentTexts.join("\n");
};
