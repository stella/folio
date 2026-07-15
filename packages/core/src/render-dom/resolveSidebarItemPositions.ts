import type { RenderedDomContext } from "./RenderedDomContext";

const MIN_CARD_GAP = 8;

export type RenderedSidebarItem = {
  id: string;
  anchorPos: number;
  anchorKey?: string;
  priority?: number;
  fixedY?: number;
  estimatedHeight?: number;
};

export type ResolvedSidebarItemPosition<T extends RenderedSidebarItem = RenderedSidebarItem> = {
  item: T;
  y: number;
};

export type ResolveSidebarItemPositionsOptions<T extends RenderedSidebarItem> = {
  items: T[];
  anchorPositions: Map<string, number>;
  renderedDomContext: RenderedDomContext | null;
  zoom: number;
  cardHeights: Map<string, number>;
  lastKnown: Map<string, number>;
};

export const resolveSidebarItemPositions = <T extends RenderedSidebarItem>({
  items,
  anchorPositions,
  renderedDomContext,
  zoom,
  cardHeights,
  lastKnown,
}: ResolveSidebarItemPositionsOptions<T>): ResolvedSidebarItemPosition<T>[] => {
  if (items.length === 0) {
    return [];
  }

  const containerOffset = renderedDomContext?.getContainerOffset();
  const positioned: Array<{ item: T; targetY: number }> = [];
  for (const item of items) {
    let targetY: number | undefined;
    if (item.fixedY !== undefined) {
      targetY = item.fixedY * zoom;
    }
    if (targetY === undefined && item.anchorKey !== undefined) {
      const anchorY = anchorPositions.get(item.anchorKey);
      if (anchorY !== undefined) {
        targetY = anchorY * zoom;
      }
    }
    if (targetY === undefined && renderedDomContext && containerOffset) {
      const rect = renderedDomContext.getRectsForRange(item.anchorPos, item.anchorPos + 1).at(0);
      if (rect) {
        targetY = (rect.y + containerOffset.y) * zoom;
      }
    }
    if (targetY === undefined) {
      targetY = lastKnown.get(item.id);
    }
    if (targetY === undefined) {
      continue;
    }
    positioned.push({ item, targetY });
    lastKnown.set(item.id, targetY);
  }

  positioned.sort((left, right) => {
    const positionDifference = left.targetY - right.targetY;
    if (positionDifference !== 0) {
      return positionDifference;
    }
    return (left.item.priority ?? 0) - (right.item.priority ?? 0);
  });

  const resolved: ResolvedSidebarItemPosition<T>[] = [];
  let previousBottom = 0;
  for (const position of positioned) {
    const height =
      cardHeights.get(position.item.id) ?? position.item.estimatedHeight ?? 80;
    const y = Math.max(position.targetY, previousBottom + MIN_CARD_GAP);
    resolved.push({ item: position.item, y });
    previousBottom = y + height;
  }
  return resolved;
};
