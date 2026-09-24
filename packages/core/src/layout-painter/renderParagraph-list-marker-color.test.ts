import { describe, expect, test } from "bun:test";

import { getRenderableListMarkerColor } from "./renderParagraph";

describe("list marker colour", () => {
  test("paints an authored colour, leaving black and auto to the canvas colour", () => {
    expect(getRenderableListMarkerColor("#2E74B5")).toBe("#2E74B5");
    expect(getRenderableListMarkerColor("#000000")).toBeUndefined();
    expect(getRenderableListMarkerColor("auto")).toBeUndefined();
    expect(getRenderableListMarkerColor(undefined)).toBeUndefined();
  });
});
