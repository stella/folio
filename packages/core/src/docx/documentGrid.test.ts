import { describe, expect, test } from "bun:test";

import { resolveDocumentGridLinePitch } from "./documentGrid";

describe("resolveDocumentGridLinePitch", () => {
  test.each(["lines", "linesAndChars", "snapToChars"] as const)(
    "activates a positive pitch for the %s grid type",
    (type) => {
      expect(resolveDocumentGridLinePitch({ type, linePitch: 360 })).toBe(360);
    },
  );

  test("keeps omitted and default grid types inactive", () => {
    expect(resolveDocumentGridLinePitch({ linePitch: 360 })).toBeUndefined();
    expect(resolveDocumentGridLinePitch({ type: "default", linePitch: 360 })).toBeUndefined();
  });

  test("rejects absent and non-positive pitches", () => {
    expect(resolveDocumentGridLinePitch({ type: "lines" })).toBeUndefined();
    expect(resolveDocumentGridLinePitch({ type: "lines", linePitch: 0 })).toBeUndefined();
    expect(resolveDocumentGridLinePitch({ type: "lines", linePitch: -1 })).toBeUndefined();
  });
});
