// Pure sidebar-card collision layout.
//
// PORT-BLOCKED: our core has no `plugin-api/resolveItemPositions` (the whole
// `plugin-api` subsystem is absent from `@stll/folio-core`). Upstream re-exported
// the core generic and its `RenderedDomContext`-based DOM anchor path. Here the
// pure-geometry algorithm is ported locally and the `RenderedDomContext` path
// (upstream step 3, DOM `getRectsForRange` lookup) is dropped: `UnifiedSidebar`
// always passed a `null` DOM context, so steps 1 (fixedY), 2 (anchorKey) and
// 4 (last-known cache) fully cover its use. Restore step 3 once a plugin-api
// (or equivalent rendered-DOM) layer lands in core.

import { MIN_CARD_GAP } from "../../utils/sidebarConstants";

/** Minimal shape the layout pass needs from a sidebar item. */
export type ResolvableSidebarItem = {
  id: string;
  anchorPos: number;
  anchorKey?: string;
  priority?: number;
  fixedY?: number;
  estimatedHeight?: number;
};

export type ResolvedPosition<T extends ResolvableSidebarItem = ResolvableSidebarItem> = {
  item: T;
  y: number;
};

/**
 * Given a list of sidebar items plus their anchor sources, return resolved Y
 * positions with collision avoidance applied. Items resolve their target Y from
 * (1) an explicit `fixedY`, (2) an `anchorKey` lookup in `anchorPositions`
 * (layout-engine coords scaled by `zoom`), or (3) the last-known cache so cards
 * don't pop out during transient layout. Overlapping cards are then pushed down
 * by their measured height + `MIN_CARD_GAP`.
 */
export function resolveItemPositions<T extends ResolvableSidebarItem>(
  items: T[],
  anchorPositions: Map<string, number>,
  zoom: number,
  cardHeights: Map<string, number>,
  lastKnown: Map<string, number>,
): ResolvedPosition<T>[] {
  if (items.length === 0) return [];

  const positioned: { item: T; targetY: number }[] = [];

  for (const item of items) {
    let y: number | undefined;

    // 1. explicit fixedY
    if (item.fixedY != null) {
      y = item.fixedY * zoom;
    }

    // 2. anchorKey lookup
    if (y == null && item.anchorKey) {
      const layoutY = anchorPositions.get(item.anchorKey);
      if (layoutY != null) y = layoutY * zoom;
    }

    // 3. cache fallback
    if (y == null) {
      const cached = lastKnown.get(item.id);
      if (cached != null) y = cached;
    }

    if (y != null) {
      positioned.push({ item, targetY: y });
      lastKnown.set(item.id, y);
    }
  }

  positioned.sort((a, b) => {
    const dy = a.targetY - b.targetY;
    if (dy !== 0) return dy;
    return (a.item.priority ?? 0) - (b.item.priority ?? 0);
  });

  const result: ResolvedPosition<T>[] = [];
  let lastBottom = 0;
  for (const pos of positioned) {
    const height = cardHeights.get(pos.item.id) ?? pos.item.estimatedHeight ?? 80;
    const y = Math.max(pos.targetY, lastBottom + MIN_CARD_GAP);
    result.push({ item: pos.item, y });
    lastBottom = y + height;
  }
  return result;
}
