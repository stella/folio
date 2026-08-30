import { describe, expect, test } from "bun:test";

import { buildFontAlternates, getFontAlternate } from "./fontAlternates";

describe("font alternates", () => {
  test("resolves document-scoped alternate names case-insensitively", () => {
    const alternates = buildFontAlternates({
      fonts: [{ name: "Brand Sans", altName: "Calibri" }],
    });

    expect(getFontAlternate(" brand SANS ", alternates)).toBe("Calibri");
  });

  test("keeps Word-compatible comma-containing alternate names intact", () => {
    const alternates = buildFontAlternates({
      fonts: [{ name: "Brand Sans", altName: "Name A, Name B" }],
    });

    expect(getFontAlternate("Brand Sans", alternates)).toBe("Name A, Name B");
  });

  test("ignores empty and self-referential alternate names", () => {
    const alternates = buildFontAlternates({
      fonts: [
        { name: "Empty", altName: "  " },
        { name: "Same", altName: " same " },
      ],
    });

    expect(alternates.size).toBe(0);
  });
});
