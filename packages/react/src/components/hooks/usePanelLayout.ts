/**
 * Measures the width the editor's panels row has and decides, through
 * {@link computePanelLayout}, how the outline and the comments are shown.
 * One ResizeObserver feeds it; the scroll gutter, the ruler offset and the
 * page centring all read the result, so they agree with what is on screen.
 */

import { type RefObject, useCallback, useEffect, useLayoutEffect, useState } from "react";

import {
  computePanelLayout,
  type PanelLayout,
  type PanelLayoutInput,
  type PanelOverlay,
} from "../panelLayout";

type UsePanelLayoutOptions = Omit<PanelLayoutInput, "availableWidth"> & {
  /** The editor's scroll container; its vertical scrollbar is not page room. */
  scrollContainerRef: RefObject<HTMLElement | null>;
};

export type PanelLayoutState = {
  /** Attach to the row the outline track and the scroll container share. */
  rowRef: (element: HTMLDivElement | null) => void;
  layout: PanelLayout;
  /** The drawer open over the page, if any. */
  overlay: PanelOverlay;
  setOverlay: (overlay: PanelOverlay) => void;
  /** The layout the same width would give with the comments open. */
  layoutWithCommentsOpen: () => PanelLayout;
};

export const usePanelLayout = ({
  pageWidth,
  outline,
  comments,
  scrollContainerRef,
}: UsePanelLayoutOptions): PanelLayoutState => {
  const [row, setRow] = useState<HTMLDivElement | null>(null);
  // Unmeasured, every panel fits; the layout effect measures before paint.
  const [availableWidth, setAvailableWidth] = useState(Number.POSITIVE_INFINITY);
  const [overlay, setOverlay] = useState<PanelOverlay>("none");

  useLayoutEffect(() => {
    if (!row) {
      return undefined;
    }
    const measure = () => {
      const scroll = scrollContainerRef.current;
      const scrollbar = scroll ? scroll.offsetWidth - scroll.clientWidth : 0;
      setAvailableWidth(Math.max(0, row.clientWidth - scrollbar));
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(row);
    // The scrollbar comes and goes with the document's height.
    if (scrollContainerRef.current) {
      observer.observe(scrollContainerRef.current);
    }
    return () => observer.disconnect();
  }, [row, scrollContainerRef]);

  const layout = computePanelLayout({ availableWidth, pageWidth, outline, comments });

  // A drawer closes once its panel has a track of its own again (or is gone).
  const outlineDrawn = layout.outline === "rail" || layout.outline === "drawer";
  const commentsDrawn = layout.comments === "drawer";
  useEffect(() => {
    if ((overlay === "outline" && !outlineDrawn) || (overlay === "comments" && !commentsDrawn)) {
      setOverlay("none");
    }
  }, [commentsDrawn, outlineDrawn, overlay]);

  const layoutWithCommentsOpen = useCallback(
    () => computePanelLayout({ availableWidth, pageWidth, outline, comments: "open" }),
    [availableWidth, outline, pageWidth],
  );

  return { rowRef: setRow, layout, overlay, setOverlay, layoutWithCommentsOpen };
};
