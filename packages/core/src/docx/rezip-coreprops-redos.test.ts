import { describe, expect, test } from "bun:test";

import { updateCoreProperties } from "./rezip";

describe("updateCoreProperties dcterms:modified", () => {
  test("rewrites an existing modified date", () => {
    const xml =
      '<cp:coreProperties><dcterms:modified xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>';
    const out = updateCoreProperties(xml, { updateModifiedDate: true });

    expect(out).not.toContain("2000-01-01");
    expect(out).toContain("<dcterms:modified");
  });

  test("does not synthesize a missing modified date", () => {
    const xml = "<cp:coreProperties><cp:revision>7</cp:revision></cp:coreProperties>";

    expect(updateCoreProperties(xml, { updateModifiedDate: true })).toBe(xml);
  });

  test("does not synthesize a missing modifier", () => {
    const xml = "<cp:coreProperties><cp:revision>7</cp:revision></cp:coreProperties>";

    expect(updateCoreProperties(xml, { modifiedBy: "Reviewer" })).toBe(xml);
  });

  test("stays linear on malformed core.xml", () => {
    const evil = "<dcterms:modified".repeat(100_000);
    const start = performance.now();
    updateCoreProperties(evil, { updateModifiedDate: true });
    expect(performance.now() - start).toBeLessThan(5000);
  });
});
