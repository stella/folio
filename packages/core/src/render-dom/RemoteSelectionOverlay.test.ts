import { describe, expect, test } from "bun:test";

import type { RenderedDomContext } from "./RenderedDomContext";
import { resolveRemoteSelectionOverlayGeometry } from "./RemoteSelectionOverlay";

describe("resolveRemoteSelectionOverlayGeometry", () => {
  test("normalizes backward ranges while keeping the caret at the remote head", () => {
    const rangeCalls: Array<[number, number]> = [];
    const pointCalls: number[] = [];
    const renderedDomContext: RenderedDomContext = {
      findElementsForRange: () => [],
      getContainerOffset: () => ({ x: 0, y: 0 }),
      getCoordinatesForPosition: (position) => {
        pointCalls.push(position);
        return { x: position, y: 4, height: 12 };
      },
      getRectsForRange: (from, to) => {
        rangeCalls.push([from, to]);
        return [{ x: from, y: 4, width: to - from, height: 12 }];
      },
    };

    const geometry = resolveRemoteSelectionOverlayGeometry(renderedDomContext, [
      { anchor: 14, clientId: 7, color: "#123456", head: 5, name: "Ada" },
    ]);

    expect(rangeCalls).toEqual([[5, 14]]);
    expect(pointCalls).toEqual([5]);
    expect(geometry).toEqual([
      {
        caret: { x: 5, y: 4, height: 12 },
        rects: [{ x: 5, y: 4, width: 9, height: 12 }],
        selection: { anchor: 14, clientId: 7, color: "#123456", head: 5, name: "Ada" },
      },
    ]);
  });
});
