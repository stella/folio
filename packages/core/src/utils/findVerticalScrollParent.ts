/**
 * Pure DOM helpers — locate the element that vertically scrolls the paginated
 * editor. Shared by scroll-to-position, arrow-key line navigation, and drag
 * auto-scroll so every framework adapter uses the same logic.
 */

/**
 * First ancestor of `el` with `overflow-y: auto|scroll` and a scrollable
 * overflow height. The walk starts at `el.parentElement` (it does not treat
 * `el` itself as the scroller). Returns `null` when no such ancestor exists
 * before `document.documentElement`.
 */
export function findVerticalScrollParent(el: HTMLElement): HTMLElement | null {
  let parent = el.parentElement;
  while (parent && parent !== document.documentElement) {
    const { overflowY } = getComputedStyle(parent);
    if (
      (overflowY === "auto" || overflowY === "scroll") &&
      parent.scrollHeight > parent.clientHeight + 1
    ) {
      return parent;
    }
    parent = parent.parentElement;
  }
  return null;
}

/**
 * Same as {@link findVerticalScrollParent} but falls back to
 * `document.documentElement` so callers always get a valid scroll target.
 */
export function findVerticalScrollParentOrRoot(el: HTMLElement): HTMLElement {
  return findVerticalScrollParent(el) ?? document.documentElement;
}
