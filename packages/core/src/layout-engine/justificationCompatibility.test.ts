import { describe, expect, test } from "bun:test";

import { resolveJustificationCompatibility } from "./justificationCompatibility";

describe("resolveJustificationCompatibility", () => {
  test.each([11, 12, 14])("uses strict justified fitting in compatibility mode %p", (mode) => {
    expect(resolveJustificationCompatibility(mode)).toEqual({ type: "legacy" });
  });

  test("reads a document without a compatibility mode with the oldest fitting rules", () => {
    expect(resolveJustificationCompatibility(undefined)).toEqual({ type: "legacy" });
  });

  test.each([15, 16])("keeps current fitting in compatibility mode %p", (mode) => {
    expect(resolveJustificationCompatibility(mode)).toBeUndefined();
  });
});
