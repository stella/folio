import { describe, expect, test } from "bun:test";

import {
  paragraphPropertySourceDelta,
  paragraphPropertySourceFingerprintFromParts,
} from "./paragraphProperties";

describe("paragraph property source delta", () => {
  test("is empty when authored pPr and paragraph-mark properties are unchanged", () => {
    const fingerprint = paragraphPropertySourceFingerprintFromParts(
      { kinsoku: false, numPr: { ilvl: 2 } },
      { runInWithNext: false },
    );

    const delta = paragraphPropertySourceDelta(fingerprint, fingerprint);

    expect(delta.isEmpty).toBe(true);
    expect(delta.changedPPrBaseKeys).toEqual([]);
    expect(delta.changedParagraphMarkKeys).toEqual([]);
    expect(delta.changedXmlGroups).toEqual([]);
  });

  test("coalesces changed modeled fields by their exact raw OOXML child", () => {
    const before = paragraphPropertySourceFingerprintFromParts(
      { spaceBefore: 120, beforeAutospacing: true, kinsoku: true },
      { runProperties: { bold: true } },
    );
    const after = paragraphPropertySourceFingerprintFromParts(
      { spaceBefore: 240, beforeAutospacing: false, kinsoku: false },
      { runProperties: { italic: true }, runInWithNext: false },
    );

    const delta = paragraphPropertySourceDelta(before, after);

    expect(delta.changedPPrBaseKeys).toEqual([
      "kinsoku",
      "spaceBefore",
      "beforeAutospacing",
    ]);
    expect(delta.changedParagraphMarkKeys).toEqual(["runProperties", "runInWithNext"]);
    expect(delta.changedXmlGroups).toEqual(["kinsoku", "spacing", "rPr"]);
    expect(delta.isEmpty).toBe(false);
    expect(Object.isFrozen(delta.changedXmlGroups)).toBe(true);
  });

  test("distinguishes explicit false and partial numbering from absence", () => {
    const empty = paragraphPropertySourceFingerprintFromParts({}, {});
    const authored = paragraphPropertySourceFingerprintFromParts(
      { suppressAutoHyphens: false, numPr: { ilvl: 0 } },
      {},
    );

    const delta = paragraphPropertySourceDelta(empty, authored);

    expect(delta.changedPPrBaseKeys).toEqual(["numPr", "suppressAutoHyphens"]);
    expect(delta.changedXmlGroups).toEqual(["numPr", "suppressAutoHyphens"]);
  });
});

