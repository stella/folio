import { describe, expect, test } from "bun:test";

import type { RenderedDomContext } from "./RenderedDomContext";
import { resolveSidebarItemPositions } from "./resolveSidebarItemPositions";

describe("resolveSidebarItemPositions", () => {
  test("uses rendered DOM geometry when semantic anchors are unavailable", () => {
    const renderedDomContext: RenderedDomContext = {
      getCoordinatesForPosition: () => null,
      findElementsForRange: () => [],
      getRectsForRange: () => [{ x: 0, y: 120, width: 10, height: 12 }],
      getContainerOffset: () => ({ x: 0, y: 20 }),
    };

    const result = resolveSidebarItemPositions({
      items: [{ id: "comment-1", anchorPos: 5 }],
      anchorPositions: new Map(),
      renderedDomContext,
      zoom: 1.5,
      cardHeights: new Map(),
      lastKnown: new Map(),
    });

    expect(result).toEqual([{ item: { id: "comment-1", anchorPos: 5 }, y: 210 }]);
  });

  test("resolves collisions deterministically by priority", () => {
    const result = resolveSidebarItemPositions({
      items: [
        { id: "later", anchorPos: 2, fixedY: 10, priority: 2, estimatedHeight: 20 },
        { id: "first", anchorPos: 1, fixedY: 10, priority: 1, estimatedHeight: 20 },
      ],
      anchorPositions: new Map(),
      renderedDomContext: null,
      zoom: 1,
      cardHeights: new Map(),
      lastKnown: new Map(),
    });

    expect(result.map(({ item, y }) => ({ id: item.id, y }))).toEqual([
      { id: "first", y: 10 },
      { id: "later", y: 38 },
    ]);
  });
});
