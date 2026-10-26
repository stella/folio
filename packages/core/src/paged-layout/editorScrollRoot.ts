/** The marked editor owns scrolling; navigation must never move host ancestors. */
import { prefersReducedMotionBehavior } from "./scrollNavigation";

const EDITOR_SCROLL_ROOT_SELECTOR = "[data-folio-scroll]";

export const getEditorScrollRoot = (
  element: HTMLElement | null,
): HTMLElement | null =>
  element?.closest<HTMLElement>(EDITOR_SCROLL_ROOT_SELECTOR) ?? null;

/** Resolve the root even when a caller holds a painted descendant or viewport. */
export const scrollEditorTo = (
  element: HTMLElement | null,
  options: ScrollToOptions,
): void => {
  getEditorScrollRoot(element)?.scrollTo(options);
};

export const scrollEditorBy = (
  element: HTMLElement | null,
  delta: number,
): void => {
  const root = getEditorScrollRoot(element);
  if (root) {
    scrollEditorTo(root, { top: root.scrollTop + delta, behavior: "instant" });
  }
};

type NearestScrollOffsetOptions = {
  targetStart: number;
  targetEnd: number;
  viewportStart: number;
  viewportEnd: number;
};

const nearestScrollOffset = ({
  targetStart,
  targetEnd,
  viewportStart,
  viewportEnd,
}: NearestScrollOffsetOptions): number => {
  const above = targetStart < viewportStart;
  const below = targetEnd > viewportEnd;
  if (above && below) return 0;
  const targetSize = targetEnd - targetStart;
  const viewportSize = viewportEnd - viewportStart;
  if (
    (above && targetSize <= viewportSize) ||
    (below && targetSize > viewportSize)
  ) {
    return targetStart - viewportStart;
  }
  if (
    (below && targetSize <= viewportSize) ||
    (above && targetSize > viewportSize)
  ) {
    return targetEnd - viewportEnd;
  }
  return 0;
};

/**
 * Farthest move, in viewport heights, that the default behavior animates. A
 * longer smooth scroll passes page shells that virtualization paints one by
 * one before discarding them, and shows blank shells until it arrives.
 */
const SMOOTH_SCROLL_MAX_VIEWPORTS = 2;

const navigationScrollBehavior = (
  distance: number,
  viewportHeight: number,
): ScrollBehavior =>
  Math.abs(distance) > viewportHeight * SMOOTH_SCROLL_MAX_VIEWPORTS
    ? "instant"
    : prefersReducedMotionBehavior();

type ScrollEditorElementOptions = {
  block?: "start" | "center" | "end" | "nearest";
  /**
   * Omitted: smooth (instant under `prefers-reduced-motion`) within two
   * viewport heights, instant beyond.
   */
  behavior?: ScrollBehavior;
  /** Space to leave between the target and the visible root edge. */
  margin?: number;
  /** Measured target range when only part of the element should be revealed. */
  rect?: DOMRectReadOnly;
};

/** Scroll only the marked root, including when the target is inside a scaled viewport. */
export const scrollEditorElementIntoView = (
  element: HTMLElement | null,
  {
    block = "center",
    behavior,
    margin = 0,
    rect,
  }: ScrollEditorElementOptions = {},
): void => {
  const root = getEditorScrollRoot(element);
  if (!root || !element) {
    return;
  }
  const targetRect = rect ?? element.getBoundingClientRect();
  const rootRect = root.getBoundingClientRect();
  const visibleTop = rootRect.top + root.clientTop;
  const targetTop = targetRect.top - visibleTop + root.scrollTop;
  const targetBottom = targetRect.bottom - visibleTop + root.scrollTop;
  let top: number;
  switch (block) {
    case "start":
      top = targetTop - margin;
      break;
    case "center":
      top = targetTop - (root.clientHeight - targetRect.height) / 2;
      break;
    case "end":
      top = targetBottom - root.clientHeight + margin;
      break;
    case "nearest":
      top =
        root.scrollTop +
        nearestScrollOffset({
          targetStart: targetTop,
          targetEnd: targetBottom,
          viewportStart: root.scrollTop + margin,
          viewportEnd: root.scrollTop + root.clientHeight - margin,
        });
      break;
    default: {
      const exhaustive: never = block;
      return exhaustive;
    }
  }
  const visibleLeft = rootRect.left + root.clientLeft;
  const left =
    root.scrollLeft +
    nearestScrollOffset({
      targetStart: targetRect.left,
      targetEnd: targetRect.right,
      viewportStart: visibleLeft,
      viewportEnd: visibleLeft + root.clientWidth,
    });
  if (top === root.scrollTop && left === root.scrollLeft) return;
  const clampedTop = Math.max(0, top);
  scrollEditorTo(root, {
    top: clampedTop,
    left: Math.max(0, left),
    behavior:
      behavior ??
      navigationScrollBehavior(clampedTop - root.scrollTop, root.clientHeight),
  });
};
