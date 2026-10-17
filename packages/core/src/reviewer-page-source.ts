import type { FolioDocumentNavigationTarget } from "./ai-edits/types";
import { resolveFolioAIBlockRange, resolveFolioAITextRange } from "./ai-edits/blockRange";
import type { FolioDocxReviewer } from "./ai-edits/headless";
import { layoutCurrentProseDocHeadless } from "./headless-layout";
import { findPageIndexContainingPmPos } from "./layout-engine/pmPageIndex";
import { hasMeasureProvider } from "./layout-engine/measure/measureProvider";
import { getPageTextFromLayout } from "./paged-layout/pageText";
import { toProseDoc } from "./prosemirror/conversion/toProseDoc";

export type FolioHeadlessPageSource = {
  pageCount(): number;
  pageText(page: number): string;
  targetPage(target: FolioDocumentNavigationTarget): number | null;
};

type CachedPages = {
  key: string;
  doc: ReturnType<typeof toProseDoc>;
  layout: ReturnType<typeof layoutCurrentProseDocHeadless>;
};

/**
 * Build a lazy page reader over a reviewer's current state. The snapshot key
 * changes for every content edit, so the next page request reprojects and
 * paginates the current document rather than serving stale pages.
 *
 * Returns `null` when no MeasureProvider is installed; callers should omit
 * page capabilities in that case rather than exposing throwing stubs.
 */
export const createFolioReviewerPageSource = (
  reviewer: FolioDocxReviewer,
): FolioHeadlessPageSource | null => {
  if (!hasMeasureProvider()) return null;

  let cached: CachedPages | null = null;
  const current = (): CachedPages => {
    const snapshot = reviewer.snapshot();
    const key = JSON.stringify(snapshot);
    if (cached?.key === key) return cached;
    const document = reviewer.toDocument();
    const doc = toProseDoc(document, {
      ...(document.package.styles ? { styles: document.package.styles } : {}),
      ...(document.package.theme ? { theme: document.package.theme } : {}),
    });
    cached = { key, doc, layout: layoutCurrentProseDocHeadless(document, doc) };
    return cached;
  };

  return {
    pageCount: () => current().layout.pages.length,
    pageText: (page) => {
      const { layout, doc } = current();
      return getPageTextFromLayout({ layout, doc, page }) ?? "";
    },
    targetPage: (target) => {
      const { layout, doc } = current();
      const snapshot = reviewer.snapshot();
      const range =
        target.type === "textRange"
          ? resolveFolioAITextRange({ range: target, doc, snapshot })
          : resolveFolioAIBlockRange({ blockId: target.blockId, doc, snapshot });
      if (range === null) return null;
      const index = findPageIndexContainingPmPos(layout, range.from);
      return index === null ? null : index + 1;
    },
  };
};
