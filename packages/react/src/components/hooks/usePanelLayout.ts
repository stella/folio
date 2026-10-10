/**
 * Measures the width the editor's panels row has and decides, through
 * {@link computePanelLayout}, how the outline and the comments are shown.
 * One ResizeObserver feeds it; the scroll gutter, the ruler offset and the
 * page centring all read the result, so they agree with what is on screen.
 */

import { useCallback, useLayoutEffect, useState } from "react";

import {
  computePanelLayout,
  type PanelLayout,
  type PanelLayoutInput,
  type PanelOverlay,
} from "@stll/folio-core/panel-layout";

type UsePanelLayoutOptions = Omit<PanelLayoutInput, "availableWidth"> & {
  /** The editor's scroll container; its vertical scrollbar is not page room. */
  scrollContainer: HTMLElement | null;
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
  scrollContainer,
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
      const scroll = scrollContainer;
      const scrollbar = scroll ? scroll.offsetWidth - scroll.clientWidth : 0;
      setAvailableWidth(Math.max(0, row.clientWidth - scrollbar));
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(row);
    // The scrollbar comes and goes with the document's height.
    if (scrollContainer) {
      observer.observe(scrollContainer);
    }
    return () => observer.disconnect();
  }, [row, scrollContainer]);

  const requestedOutline = overlay === "outline" && outline !== "absent" ? "expanded" : outline;
  const layout = computePanelLayout({
    availableWidth,
    pageWidth,
    outline: requestedOutline,
    comments,
  });

  // An overlay closes if its panel is no longer available.
  const outlineDrawn = layout.outline !== "none" && layout.outline !== "column";
  const commentsDrawn = layout.comments === "drawer";
  if ((overlay === "outline" && !outlineDrawn) || (overlay === "comments" && !commentsDrawn)) {
    setOverlay("none");
  }

  const layoutWithCommentsOpen = useCallback(
    () =>
      computePanelLayout({
        availableWidth,
        pageWidth,
        outline: requestedOutline,
        comments: "open",
      }),
    [availableWidth, pageWidth, requestedOutline],
  );

  return { rowRef: setRow, layout, overlay, setOverlay, layoutWithCommentsOpen };
};
