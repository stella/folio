import {
  type KeyboardEvent,
  type RefObject,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from "react";

import { cn } from "../../lib/utils";
import type { FolioOutlineRailProps, OutlineItem } from "../folio-ui";

/**
 * Built-in, dependency-light OutlineRail used when a consumer does not inject
 * one. The editor mounts it inside the outline column, drawer or rail, so it
 * fills its parent:
 *
 * - `panel`: a list of headings nested by level, one line each (the full text
 *   in a tooltip), the active heading marked and kept in view.
 * - `rail`: one tick per heading, placed where the heading falls in the
 *   document and indented by level; hovering or focusing a tick names it.
 *
 * Both are one tab stop: the arrow keys, Home and End move between headings
 * (a roving tabindex) and Enter jumps.
 */
export function DefaultOutlineRail({
  items,
  scrollContainerRef,
  resolvePct,
  onJump,
  activeId,
  ariaLabel = "Outline",
  presentation = "panel",
}: FolioOutlineRailProps) {
  const listRef = useRef<HTMLOListElement>(null);
  const activeIndex = items.findIndex((item) => item.id === activeId);
  const [focusIndex, setFocusIndex] = useState(Math.max(0, activeIndex));
  const lastIndex = items.length - 1;
  const clampedFocusIndex = Math.max(0, Math.min(lastIndex, focusIndex));
  let minLevel = Infinity;
  for (const item of items) {
    minLevel = Math.min(minLevel, item.level);
  }

  // The tab stop follows the active heading until the user moves it.
  const focusWithin = useRef(false);
  useEffect(() => {
    if (!focusWithin.current && activeIndex >= 0) {
      setFocusIndex(activeIndex);
    }
  }, [activeIndex]);

  // Keep the active heading in view without scrolling any ancestor (a plain
  // `scrollIntoView` would also scroll the editor behind the panel).
  useEffect(() => {
    const list = listRef.current;
    if (!list || presentation !== "panel" || activeIndex < 0) {
      return;
    }
    const entry = list.children.item(activeIndex);
    if (!(entry instanceof HTMLElement)) {
      return;
    }
    const top = entry.offsetTop;
    const bottom = top + entry.offsetHeight;
    if (top < list.scrollTop) {
      list.scrollTop = top;
    } else if (bottom > list.scrollTop + list.clientHeight) {
      list.scrollTop = bottom - list.clientHeight;
    }
  }, [activeIndex, presentation]);

  const jump = useCallback(
    (item: OutlineItem) => {
      const container = scrollContainerRef.current;
      if (container) {
        onJump(item.id, container);
      }
    },
    [onJump, scrollContainerRef],
  );

  const handleKeyDown = useCallback(
    (event: KeyboardEvent<HTMLOListElement>) => {
      const next = (() => {
        switch (event.key) {
          case "ArrowDown":
            return Math.min(lastIndex, clampedFocusIndex + 1);
          case "ArrowUp":
            return Math.max(0, clampedFocusIndex - 1);
          case "Home":
            return 0;
          case "End":
            return lastIndex;
          default:
            return null;
        }
      })();
      if (next === null) {
        return;
      }
      event.preventDefault();
      setFocusIndex(next);
      const buttons = listRef.current?.querySelectorAll<HTMLButtonElement>("button");
      if (buttons) {
        [...buttons].at(next)?.focus();
      }
    },
    [clampedFocusIndex, lastIndex],
  );

  const tickTops = useRailTickTops({
    enabled: presentation === "rail",
    items,
    listRef,
    resolvePct,
    scrollContainerRef,
  });

  if (items.length < 2) {
    return null;
  }

  const rail = presentation === "rail";
  return (
    <ol
      ref={listRef}
      aria-label={ariaLabel}
      className={rail ? "folio-outline-ticks" : "folio-outline-list"}
      onKeyDown={handleKeyDown}
      onFocus={() => {
        focusWithin.current = true;
      }}
      onBlur={(event) => {
        const next = event.relatedTarget;
        if (!(next instanceof Node && event.currentTarget.contains(next))) {
          focusWithin.current = false;
        }
      }}
    >
      {items.map((item, index) => {
        const isActive = item.id === activeId;
        const depth = Math.min(4, item.level - minLevel);
        return (
          <li
            key={item.id}
            className={rail ? "folio-outline-tick-slot" : undefined}
            style={rail ? { top: tickTops[index] ?? 0 } : undefined}
          >
            <button
              aria-current={isActive ? "true" : undefined}
              aria-label={rail ? item.label : undefined}
              className={cn(
                rail ? "folio-outline-tick" : "folio-outline-item",
                isActive && (rail ? "folio-outline-tick--active" : "folio-outline-item--active"),
              )}
              data-depth={depth}
              onClick={() => jump(item)}
              onFocus={() => setFocusIndex(index)}
              tabIndex={index === clampedFocusIndex ? 0 : -1}
              title={rail ? undefined : item.label}
              type="button"
            >
              {rail ? (
                <>
                  <span aria-hidden="true" className="folio-outline-tick-mark" />
                  <span aria-hidden="true" className="folio-outline-tick-label">
                    {item.label}
                  </span>
                </>
              ) : (
                <span className="folio-outline-item-label">{item.label}</span>
              )}
            </button>
          </li>
        );
      })}
    </ol>
  );
}

/** Height of one tick's hit area; ticks never overlap by less than this. */
const TICK_PITCH_PX = 12;

type RailTickTopsOptions = {
  enabled: boolean;
  items: OutlineItem[];
  listRef: RefObject<HTMLOListElement | null>;
  resolvePct: FolioOutlineRailProps["resolvePct"];
  scrollContainerRef: FolioOutlineRailProps["scrollContainerRef"];
};

/**
 * Pixel offsets for each tick: where its heading falls in the document,
 * pushed apart so no two ticks share a hit area, and pulled back inside the
 * rail. When the rail is too short for every tick, they are spread evenly.
 */
function useRailTickTops({
  enabled,
  items,
  listRef,
  resolvePct,
  scrollContainerRef,
}: RailTickTopsOptions): number[] {
  const [height, setHeight] = useState(0);
  useLayoutEffect(() => {
    const list = listRef.current;
    if (!enabled || !list) {
      return undefined;
    }
    setHeight(list.clientHeight);
    const observer = new ResizeObserver(() => setHeight(list.clientHeight));
    observer.observe(list);
    return () => observer.disconnect();
  }, [enabled, listRef]);

  if (!enabled || height === 0) {
    return [];
  }
  const container = scrollContainerRef.current;
  const span = Math.max(0, height - TICK_PITCH_PX);
  if (items.length * TICK_PITCH_PX > height) {
    const step = items.length > 1 ? span / (items.length - 1) : 0;
    return items.map((_, index) => index * step);
  }
  const tops = items.map((item, index) => {
    const pct = container ? resolvePct(item.id, container) : null;
    const fraction = pct === null ? index / Math.max(1, items.length - 1) : pct / 100;
    return fraction * span;
  });
  for (let index = 1; index < tops.length; index++) {
    tops[index] = Math.max(tops[index] ?? 0, (tops[index - 1] ?? 0) + TICK_PITCH_PX);
  }
  for (let index = tops.length - 1; index >= 0; index--) {
    const limit = index === tops.length - 1 ? span : (tops[index + 1] ?? span) - TICK_PITCH_PX;
    tops[index] = Math.min(tops[index] ?? 0, limit);
  }
  return tops;
}
