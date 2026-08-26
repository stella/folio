import { resolveSequentialBlockAnchor } from "../ai-edits/blockRange";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { findPageShellForPmPos, PAINTER_PAINTED_EVENT } from "../layout-painter/renderPage";

/** Painted geometry for one AI-snapshot block in scroll-content coordinates. */
export type BlockRect = {
  blockId: string;
  /** One-based painted page number. */
  page: number;
  /** Pixels from the scroll root's content origin. */
  top: number;
  /** Pixels from the scroll root's content origin. */
  left: number;
  width: number;
  height: number;
};

export type ReadBlockRectsOptions = {
  blockIds: readonly string[];
  snapshot: FolioAIEditSnapshot;
  pagesContainer: HTMLElement;
  scrollRoot: HTMLElement;
};

type RequestedBlock = {
  blockId: string;
  pmStart: number;
};

const PAINTED_BLOCK_SELECTOR = "[data-block-id][data-pm-start]";
const PAGE_SELECTOR = "[data-page-number]";

const pageNumberOf = (page: HTMLElement): number | null => {
  const raw = page.dataset["pageNumber"];
  const pageNumber = raw === undefined ? Number.NaN : Number(raw);
  return Number.isInteger(pageNumber) && pageNumber > 0 ? pageNumber : null;
};

const requestedBlocks = (
  blockIds: readonly string[],
  snapshot: FolioAIEditSnapshot,
): RequestedBlock[] => {
  const uniqueIds = new Set(blockIds);
  const requested: RequestedBlock[] = [];
  for (const blockId of uniqueIds) {
    // Same resolution as `resolveFolioAIBlockRange`: an own-property anchor,
    // else a positional `seq-NNNN` id minted for a paraId-less paragraph.
    const anchor = Object.hasOwn(snapshot.anchors, blockId)
      ? snapshot.anchors[blockId]
      : resolveSequentialBlockAnchor(blockId, snapshot);
    if (anchor) {
      requested.push({ blockId, pmStart: anchor.from });
    }
  }
  return requested;
};

const paintedBlocksByPmStart = (page: HTMLElement): ReadonlyMap<number, HTMLElement> => {
  const blocks = new Map<number, HTMLElement>();
  for (const element of page.querySelectorAll<HTMLElement>(PAINTED_BLOCK_SELECTOR)) {
    const pmStart = Number(element.dataset["pmStart"]);
    if (Number.isFinite(pmStart) && !blocks.has(pmStart)) {
      blocks.set(pmStart, element);
    }
  }
  return blocks;
};

const groupVirtualizedRequestsByPage = (
  requests: readonly RequestedBlock[],
  pagesContainer: HTMLElement,
): ReadonlyMap<HTMLElement, RequestedBlock[]> | null => {
  const byPage = new Map<HTMLElement, RequestedBlock[]>();
  for (const request of requests) {
    const page = findPageShellForPmPos(pagesContainer, request.pmStart)?.element;
    if (!page) {
      return null;
    }
    const pageRequests = byPage.get(page);
    if (pageRequests) {
      pageRequests.push(request);
    } else {
      byPage.set(page, [request]);
    }
  }
  return byPage;
};

const findPaintedElements = (
  requests: readonly RequestedBlock[],
  pagesContainer: HTMLElement,
): ReadonlyMap<string, { element: HTMLElement; page: number }> => {
  const found = new Map<string, { element: HTMLElement; page: number }>();
  const virtualizedRequests = groupVirtualizedRequestsByPage(requests, pagesContainer);

  if (virtualizedRequests) {
    for (const [page, pageRequests] of virtualizedRequests) {
      const pageNumber = pageNumberOf(page);
      if (pageNumber === null) {
        continue;
      }
      const blocks = paintedBlocksByPmStart(page);
      for (const request of pageRequests) {
        const element = blocks.get(request.pmStart);
        if (element) {
          found.set(request.blockId, { element, page: pageNumber });
        }
      }
    }
    return found;
  }

  const requestsByPmStart = new Map(requests.map((request) => [request.pmStart, request]));
  for (const page of pagesContainer.querySelectorAll<HTMLElement>(PAGE_SELECTOR)) {
    const pageNumber = pageNumberOf(page);
    if (pageNumber === null) {
      continue;
    }
    const blocks = paintedBlocksByPmStart(page);
    for (const [pmStart, request] of requestsByPmStart) {
      const element = blocks.get(pmStart);
      if (element) {
        found.set(request.blockId, { element, page: pageNumber });
        requestsByPmStart.delete(pmStart);
      }
    }
    if (requestsByPmStart.size === 0) {
      break;
    }
  }
  return found;
};

/**
 * Measure painted snapshot blocks relative to one editor's scroll root.
 *
 * Missing snapshot ids and blocks on unpainted virtual pages are omitted. Each
 * requested block is measured at most once; painted elements are collected in
 * one DOM query per relevant page.
 */
export const readBlockRects = ({
  blockIds,
  snapshot,
  pagesContainer,
  scrollRoot,
}: ReadBlockRectsOptions): ReadonlyMap<string, BlockRect> => {
  const requests = requestedBlocks(blockIds, snapshot);
  const paintedElements = findPaintedElements(requests, pagesContainer);
  if (paintedElements.size === 0) {
    return new Map();
  }

  const rootRect = scrollRoot.getBoundingClientRect();
  const rects = new Map<string, BlockRect>();
  for (const request of requests) {
    const painted = paintedElements.get(request.blockId);
    if (!painted) {
      continue;
    }
    const rect = painted.element.getBoundingClientRect();
    rects.set(request.blockId, {
      blockId: request.blockId,
      page: painted.page,
      top: rect.top - rootRect.top - scrollRoot.clientTop + scrollRoot.scrollTop,
      left: rect.left - rootRect.left - scrollRoot.clientLeft + scrollRoot.scrollLeft,
      width: rect.width,
      height: rect.height,
    });
  }
  return rects;
};

/** Subscribe to instance-scoped painted-page changes. */
export const onPaintedLayoutChange = (
  pagesContainer: HTMLElement,
  listener: () => void,
): (() => void) => {
  pagesContainer.addEventListener(PAINTER_PAINTED_EVENT, listener);
  return () => {
    pagesContainer.removeEventListener(PAINTER_PAINTED_EVENT, listener);
  };
};
