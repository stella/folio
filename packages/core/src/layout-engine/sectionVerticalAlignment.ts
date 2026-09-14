import type { VerticalAlign } from "@stll/docx-core/model";

import type { Fragment, Layout, Page } from "./types";

const isFlowFragment = (fragment: Fragment): boolean => {
  switch (fragment.kind) {
    case "paragraph":
      return true;
    case "table":
      return !fragment.isFloating;
    case "image":
      return !fragment.isAnchored;
    case "textBox":
      return !fragment.isPositioned;
    default:
      fragment satisfies never;
      return false;
  }
};

const alignedOffset = (alignment: VerticalAlign, availableHeight: number): number => {
  switch (alignment) {
    case "center":
      return availableHeight / 2;
    case "bottom":
      return availableHeight;
    case "top":
    case "both":
      return 0;
    default:
      alignment satisfies never;
      return 0;
  }
};

const alignPage = (page: Page, alignment: VerticalAlign | undefined): Page => {
  if (alignment !== "center" && alignment !== "bottom") {
    return page;
  }

  const flow = page.fragments.filter(isFlowFragment);
  if (flow.length === 0) {
    return page;
  }

  const contentBottom = page.size.h - page.margins.bottom - (page.footnoteReservedHeight ?? 0);
  const occupiedBottom = Math.max(...flow.map((fragment) => fragment.y + fragment.height));
  const offset = alignedOffset(alignment, Math.max(0, contentBottom - occupiedBottom));
  if (offset === 0) {
    return page;
  }

  const flowSet = new Set(flow);
  return {
    ...page,
    fragments: page.fragments.map((fragment) =>
      flowSet.has(fragment) ? { ...fragment, y: fragment.y + offset } : fragment,
    ),
  };
};

/** Align each page's body inside the authored section content frame. */
export const applySectionVerticalAlignment = (
  layout: Layout,
  alignments: readonly (VerticalAlign | undefined)[] | undefined,
): Layout => {
  if (alignments === undefined) {
    return layout;
  }

  let changed = false;
  const pages = layout.pages.map((page) => {
    const aligned = alignPage(page, alignments[page.sectionIndex ?? 0]);
    changed ||= aligned !== page;
    return aligned;
  });
  return changed ? { ...layout, pages } : layout;
};
