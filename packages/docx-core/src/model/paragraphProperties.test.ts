import { describe, expect, test } from "bun:test";

import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  PARAGRAPH_PROPERTY_SOURCE_XML_GROUP,
  paragraphPropertySourceDelta,
  paragraphPropertySourceFingerprintFromParts,
} from "./paragraphProperties";
import { COLOR_VALUE_PROPERTY_DESCRIPTORS, SHADING_PROPERTY_DESCRIPTORS } from "./colors";
import {
  TEXT_FORMATTING_FONT_FAMILY_FIELD_DESCRIPTORS,
  TEXT_FORMATTING_LANGUAGE_FIELD_DESCRIPTORS,
  TEXT_FORMATTING_PROPERTY_DESCRIPTORS,
  TEXT_FORMATTING_UNDERLINE_FIELD_DESCRIPTORS,
} from "./formatting";

describe("paragraph property source delta", () => {
  test("published formatting descriptors are recursively immutable", () => {
    for (const descriptor of [
      COLOR_VALUE_PROPERTY_DESCRIPTORS,
      SHADING_PROPERTY_DESCRIPTORS,
      TEXT_FORMATTING_PROPERTY_DESCRIPTORS,
      TEXT_FORMATTING_FONT_FAMILY_FIELD_DESCRIPTORS,
      TEXT_FORMATTING_LANGUAGE_FIELD_DESCRIPTORS,
      TEXT_FORMATTING_UNDERLINE_FIELD_DESCRIPTORS,
      PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
    ]) {
      expect(Object.isFrozen(descriptor)).toBe(true);
      for (const entry of Object.values(descriptor)) {
        expect(Object.isFrozen(entry)).toBe(true);
      }
    }
    expect(Object.isFrozen(PARAGRAPH_PROPERTY_SOURCE_XML_GROUP)).toBe(true);
  });

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

    expect(delta.changedPPrBaseKeys).toEqual(["kinsoku", "spaceBefore", "beforeAutospacing"]);
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
