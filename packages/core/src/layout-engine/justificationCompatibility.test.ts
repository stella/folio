import { describe, expect, test } from "bun:test";

import { resolveJustificationCompatibility } from "./justificationCompatibility";

describe("resolveJustificationCompatibility", () => {
  test("uses strict justified fitting through compatibility mode 14", () => {
    expect(resolveJustificationCompatibility(14)).toEqual({ type: "legacy" });
  });

  test("keeps modern and unspecified documents on current fitting", () => {
    expect(resolveJustificationCompatibility(15)).toBeUndefined();
    expect(resolveJustificationCompatibility(undefined)).toBeUndefined();
  });
});
