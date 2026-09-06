import { describe, expect, test } from "bun:test";

import { resolveTableIndentCompatibility } from "./tableIndentCompatibility";

describe("resolveTableIndentCompatibility", () => {
  test.each([undefined, 11, 12, 14])("reads compatibilityMode %p as a text-edge indent", (mode) => {
    expect(resolveTableIndentCompatibility(mode)).toEqual({ type: "legacy" });
  });

  test.each([15, 16])("reads compatibilityMode %p as a border-edge indent", (mode) => {
    expect(resolveTableIndentCompatibility(mode)).toBeUndefined();
  });
});
