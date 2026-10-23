import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import type { Dispatch, RefObject, SetStateAction } from "react";

import { scrollEditorTo } from "@stll/folio-core/paged-layout/editorScrollRoot";
import { fitWidthZoom } from "@stll/folio-core/paged-layout/fitWidthZoom";
import { computeScrollPageInfo } from "@stll/folio-core/paged-layout/scrollPageInfo";
import type { ScrollPageInfo } from "@stll/folio-core/paged-layout/scrollPageInfo";
import {
  getScrollTopForZoomAnchor,
  getViewportCenterZoomAnchorForZoomChange,
} from "@stll/folio-core/paged-layout/zoomScrollAnchor";
import type { ViewportCenterZoomAnchor } from "@stll/folio-core/paged-layout/zoomScrollAnchor";

import {
  DEFAULT_PAGE_GAP,
  VIEWPORT_PADDING_TOP,
} from "../../paged-editor/PagedEditor";
import type { PagedEditorRef } from "../../paged-editor/PagedEditor";

const PAGE_INFO_FADE_MS = 600;

export type UseZoomAndPageInfoArgs = {
  scrollContainerRef: RefObject<HTMLDivElement | null>;
  scrollContainer: HTMLDivElement | null;
  pagedEditorRef: RefObject<PagedEditorRef | null>;
  /**
   * Initial zoom on mount as a number (1 = 100%), or `"fit-width"` to size the
   * page to the scroll container's width and keep it fit as the container
   * resizes. A later imperative `setZoom`/toolbar change overrides the fit.
   */
  initialZoom: number | "fit-width";
};

export type UseZoomAndPageInfoReturn = {
  zoom: number;
  /** Stable ref that mirrors `zoom` for use in callbacks that need the
   *  current zoom without re-binding (e.g., scroll handlers). */
  zoomRef: RefObject<number>;
  /**
   * Set a new zoom level. Captures the viewport center so the editor
   * re-anchors after the layout reflows.
   */
  setZoomWithViewportAnchor: (zoom: number) => void;
  scrollPageInfo: ScrollPageInfo;
  setScrollPageInfo: Dispatch<SetStateAction<ScrollPageInfo>>;
  /** Recompute current/total page indicators from the live scroll position. */
  updateScrollPageInfo: (scrollContainer: HTMLDivElement) => void;
  /** Hide the page indicator overlay after a short delay. */
  scheduleScrollPageInfoFade: () => void;
};

export function useZoomAndPageInfo({
  scrollContainerRef,
  scrollContainer: scrollContainerEl,
  pagedEditorRef,
  initialZoom,
}: UseZoomAndPageInfoArgs): UseZoomAndPageInfoReturn {
  const fitToWidth = initialZoom === "fit-width";
  const [zoom, setZoom] = useState(
    typeof initialZoom === "number" ? initialZoom : 1,
  );
  const zoomRef = useRef(zoom);
  const pendingZoomAnchorRef = useRef<ViewportCenterZoomAnchor | null>(null);

  const [scrollPageInfo, setScrollPageInfo] = useState<ScrollPageInfo>({
    currentPage: 1,
    totalPages: 1,
    visible: false,
  });
  const scrollFadeTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const scheduleScrollPageInfoFade = useCallback(() => {
    if (scrollFadeTimerRef.current) {
      clearTimeout(scrollFadeTimerRef.current);
    }
    scrollFadeTimerRef.current = setTimeout(() => {
      setScrollPageInfo((prev) => ({ ...prev, visible: false }));
    }, PAGE_INFO_FADE_MS);
  }, []);

  const updateScrollPageInfo = useCallback(
    (scrollContainer: HTMLDivElement) => {
      const layout = pagedEditorRef.current?.getLayout();
      if (!layout) {
        return;
      }

      const info = computeScrollPageInfo({
        pageHeights: layout.pages.map((page) => page.size.h),
        scrollTop: scrollContainer.scrollTop,
        clientHeight: scrollContainer.clientHeight,
        zoom: zoomRef.current,
        viewportPaddingTop: VIEWPORT_PADDING_TOP,
        pageGap: DEFAULT_PAGE_GAP,
      });
      if (!info) {
        return;
      }

      setScrollPageInfo({ ...info, visible: true });
    },
    [pagedEditorRef],
  );

  const setZoomWithViewportAnchor = useCallback(
    (nextZoom: number) => {
      const currentZoom = zoomRef.current;
      const scrollContainer = scrollContainerRef.current;
      const previousAnchor = pendingZoomAnchorRef.current;

      pendingZoomAnchorRef.current = scrollContainer
        ? getViewportCenterZoomAnchorForZoomChange({
            clientHeight: scrollContainer.clientHeight,
            currentZoom,
            nextZoom,
            pendingAnchor: previousAnchor,
            scrollTop: scrollContainer.scrollTop,
          })
        : null;

      if (currentZoom === nextZoom) {
        return;
      }
      zoomRef.current = nextZoom;
      setZoom(nextZoom);
    },
    [scrollContainerRef],
  );

  useLayoutEffect(() => {
    zoomRef.current = zoom;
  }, [zoom]);

  useLayoutEffect(() => {
    const anchor = pendingZoomAnchorRef.current;
    if (!anchor) {
      return;
    }
    pendingZoomAnchorRef.current = null;
    if (anchor.zoom === zoom) {
      return;
    }
    const scrollContainer = scrollContainerRef.current;
    if (!scrollContainer) {
      return;
    }
    const nextScrollTop = getScrollTopForZoomAnchor(anchor, zoom);
    scrollEditorTo(scrollContainer, {
      top: nextScrollTop,
      behavior: "instant",
    });
    updateScrollPageInfo(scrollContainer);
    scheduleScrollPageInfoFade();
  }, [
    zoom,
    scheduleScrollPageInfoFade,
    updateScrollPageInfo,
    scrollContainerRef,
  ]);

  // Subscribe when the callback ref publishes the committed scroll container.
  useEffect(() => {
    if (!scrollContainerEl) {
      return;
    }
    const handleScroll = () => {
      updateScrollPageInfo(scrollContainerEl);
      scheduleScrollPageInfoFade();
    };
    scrollContainerEl.addEventListener("scroll", handleScroll, {
      passive: true,
    });
    return () => {
      scrollContainerEl.removeEventListener("scroll", handleScroll);
      if (scrollFadeTimerRef.current) {
        clearTimeout(scrollFadeTimerRef.current);
      }
    };
  }, [scrollContainerEl, scheduleScrollPageInfoFade, updateScrollPageInfo]);

  // fit-width: scale the page so its width fills the scroll container, and keep
  // it fit as the container resizes. The page's natural width comes from the
  // laid-out geometry (`layout.pages[0].size.w`), so the scaled page never has
  // to be measured in the DOM; a rAF retry covers the window before the first
  // layout completes. A manual zoom afterwards wins (this only re-fires on a
  // container resize), matching the documented "override the fit" behaviour.
  // The fit basis reserves the widest scrollbar gutter seen on this container,
  // so a gutter the fit itself toggles cannot feed back into the next fit.
  useEffect(() => {
    if (!fitToWidth || !scrollContainerEl) {
      return undefined;
    }
    let rafId = 0;
    let reservedEdge = 0;
    const applyFit = () => {
      const pageWidth =
        pagedEditorRef.current?.getLayout()?.pages[0]?.size.w ?? 0;
      if (pageWidth <= 0) {
        rafId = requestAnimationFrame(applyFit); // layout not ready yet
        return;
      }
      const fit = fitWidthZoom({
        offsetWidth: scrollContainerEl.offsetWidth,
        clientWidth: scrollContainerEl.clientWidth,
        pageWidth,
        reservedEdge,
      });
      reservedEdge = fit.reservedEdge;
      // Guard against a redundant zoom set (which would re-anchor the scroll)
      // and sub-percent jitter.
      if (fit.zoom !== null && Math.abs(fit.zoom - zoomRef.current) > 0.005) {
        setZoomWithViewportAnchor(fit.zoom);
      }
    };
    const observer = new ResizeObserver(() => {
      cancelAnimationFrame(rafId);
      rafId = requestAnimationFrame(applyFit);
    });
    observer.observe(scrollContainerEl);
    applyFit();
    return () => {
      cancelAnimationFrame(rafId);
      observer.disconnect();
    };
  }, [
    fitToWidth,
    scrollContainerEl,
    pagedEditorRef,
    setZoomWithViewportAnchor,
  ]);

  return {
    zoom,
    zoomRef,
    setZoomWithViewportAnchor,
    scrollPageInfo,
    setScrollPageInfo,
    updateScrollPageInfo,
    scheduleScrollPageInfoFade,
  };
}
