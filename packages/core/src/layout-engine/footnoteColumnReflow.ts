import type { Fragment, Layout } from "./types";

function isFlowFragment(fragment: Fragment): boolean {
  switch (fragment.kind) {
    case "paragraph":
      return true;
    case "table":
      return !fragment.isFloating;
    case "image":
      return !fragment.isAnchored;
    case "textBox":
      return !fragment.isPositioned;
    default: {
      const exhaustive: never = fragment;
      return exhaustive;
    }
  }
}

function hasFootnoteColumnOverlap(layout: Layout): boolean {
  for (const page of layout.pages) {
    if ((page.columns?.count ?? 1) < 2 || !page.footnoteReservedHeight) {
      continue;
    }
    const bodyBottom = page.size.h - page.margins.bottom - page.footnoteReservedHeight;
    if (
      page.fragments.some(
        (fragment) => isFlowFragment(fragment) && fragment.y + fragment.height > bodyBottom,
      )
    ) {
      return true;
    }
  }
  return false;
}

function growFootnoteReserveFloors(
  layout: Layout,
  initialReserveFloors: ReadonlyMap<number, number> | undefined,
): Map<number, number> {
  const next = new Map<number, number>();
  for (const page of layout.pages) {
    // In the dynamic footnote path, a retry floor is owned by the IDs on its
    // page. A caller floor is authoritative while those IDs remain, and the
    // current observed reservation can only increase it. Entries without
    // current IDs are intentionally omitted so a moved note cannot strand a
    // floor on its former page. Static floors remain supported by
    // layoutDocument when no dynamic footnote map is supplied.
    if ((page.footnoteIds?.length ?? 0) === 0) {
      continue;
    }
    const observed = page.footnoteReservedHeight ?? 0;
    const callerFloor = initialReserveFloors?.get(page.number) ?? 0;
    const floor = Math.max(callerFloor, observed);
    if (floor > 0) {
      next.set(page.number, floor);
    }
  }
  return next;
}

function reserveFloorFingerprint(floors: ReadonlyMap<number, number>): string {
  return JSON.stringify(
    [...floors.entries()].sort(([leftPage], [rightPage]) => leftPage - rightPage),
  );
}

function clearOrphanedFootnoteReservations(layout: Layout): Layout {
  let pagesChanged = false;
  const pages = layout.pages.map((page) => {
    if ((page.footnoteIds?.length ?? 0) > 0 || page.footnoteReservedHeight === undefined) {
      return page;
    }
    pagesChanged = true;
    const cleanedPage = { ...page };
    delete cleanedPage.footnoteReservedHeight;
    return cleanedPage;
  });

  if (!pagesChanged) {
    return layout;
  }
  return { ...layout, pages };
}

function getPassBudget(layout: Layout): number {
  const fragmentCount = layout.pages.reduce((count, page) => count + page.fragments.length, 0);
  const footnoteCount = new Set(layout.pages.flatMap((page) => page.footnoteIds ?? [])).size;

  // Every retry must change a finite page/fragment assignment or terminate.
  // The budget scales with the document state rather than imposing a small
  // constant that rejects long but valid cascades.
  return layout.pages.length + fragmentCount + footnoteCount + 1;
}

type ReflowFootnoteColumnsOptions = {
  initialLayout: Layout;
  initialReserveFloors?: ReadonlyMap<number, number>;
  runLayout: (reserveFloors: Map<number, number>) => Layout;
};

/**
 * Re-run a multi-column layout until its body flow clears the shared footnote band.
 * Reserve floors follow the current page-to-footnote assignment. Repeated floor states
 * and persistent overlap use a stable fallback layout instead of throwing.
 */
export function reflowFootnoteColumns({
  initialLayout,
  initialReserveFloors,
  runLayout,
}: ReflowFootnoteColumnsOptions): Layout {
  let layout = initialLayout;
  if (!hasFootnoteColumnOverlap(layout)) {
    return clearOrphanedFootnoteReservations(layout);
  }

  const seenReserveStates = new Set<string>();
  const passBudget = getPassBudget(layout);
  for (let pass = 0; pass < passBudget; pass += 1) {
    const reserveFloors = growFootnoteReserveFloors(layout, initialReserveFloors);
    const reserveState = reserveFloorFingerprint(reserveFloors);
    if (seenReserveStates.has(reserveState)) {
      // A deterministic cycle has no new geometry to discover. Preserve the
      // latest layout as an explicit stable fallback instead of throwing from
      // an untrusted document's layout path.
      return clearOrphanedFootnoteReservations(layout);
    }
    seenReserveStates.add(reserveState);

    layout = runLayout(reserveFloors);
    if (!hasFootnoteColumnOverlap(layout)) {
      return clearOrphanedFootnoteReservations(layout);
    }
  }

  // The document-state budget was exhausted without a non-overlapping fixed
  // point. Keep the last deterministic layout; callers can still render and
  // inspect the document, including the footnote IDs assigned to each page.
  return clearOrphanedFootnoteReservations(layout);
}
