/**
 * The heading the reader is in: the last heading whose paragraph starts above
 * the middle of the scroll viewport, tracked as the document scrolls. Shared
 * by every place the outline is drawn (column, rail, drawer), so a drawer
 * opens on the same heading the rail marks.
 */

import { type RefObject, useCallback, useEffect, useRef, useState } from "react";

import { findBodyPmAnchors } from "@stll/folio-core/layout-bridge/dom/findBodyPmSpans";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";

/** How long a jump holds the active heading before scrolling may move it. */
const MANUAL_ACTIVE_LOCK_MS = 900;

/** Outline item id for a heading: its ProseMirror position. */
export const headingId = (heading: HeadingInfo) => String(heading.pmPos);

type ActiveHeading = {
  activeId: string | null;
  /** Mark `id` active now, holding it while the jump's scroll settles. */
  markJumped: (id: string) => void;
};

export const useActiveHeading = (
  scrollContainerRef: RefObject<HTMLElement | null>,
  headings: readonly HeadingInfo[],
): ActiveHeading => {
  const [activeId, setActiveId] = useState<string | null>(null);
  // Suppress the scroll-driven detector briefly after a jump so the in-flight
  // smooth scroll doesn't revert the highlight to the prior heading.
  const manualLockUntil = useRef(0);

  useEffect(() => {
    const container = scrollContainerRef.current;
    if (!container || headings.length === 0) {
      return undefined;
    }
    const compute = () => {
      if (Date.now() < manualLockUntil.current) {
        return;
      }
      const anchors = findBodyPmAnchors(container);
      if (anchors.length === 0) {
        return;
      }
      const containerTop = container.getBoundingClientRect().top;
      const threshold = container.scrollTop + container.clientHeight / 2;
      const offsets: { pm: number; top: number }[] = [];
      for (const el of anchors) {
        const pm = Number(el.dataset["pmStart"]);
        if (!Number.isFinite(pm)) {
          continue;
        }
        offsets.push({
          pm,
          top: el.getBoundingClientRect().top - containerTop + container.scrollTop,
        });
      }
      offsets.sort((a, b) => a.pm - b.pm);

      let next: string | null = null;
      for (const heading of headings) {
        const candidate = offsets.find((offset) => offset.pm >= heading.pmPos);
        if (candidate && candidate.top <= threshold) {
          next = headingId(heading);
        }
      }
      setActiveId(next);
    };

    let raf = 0;
    const onScroll = () => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(compute);
    };
    container.addEventListener("scroll", onScroll, { passive: true });
    raf = requestAnimationFrame(compute);
    return () => {
      cancelAnimationFrame(raf);
      container.removeEventListener("scroll", onScroll);
    };
  }, [scrollContainerRef, headings]);

  const markJumped = useCallback((id: string) => {
    setActiveId(id);
    manualLockUntil.current = Date.now() + MANUAL_ACTIVE_LOCK_MS;
  }, []);

  return { activeId, markJumped };
};
