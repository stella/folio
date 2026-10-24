import { expect, test } from "bun:test";

import type { FontInfo } from "@stll/docx-core/model";

import { hashParagraphBlock } from "../layout-engine/measure/cache";
import type { ParagraphBlock } from "../layout-engine/types";
import { buildDocumentFontFacts } from "./documentFontFacts";
import { embeddedFontFingerprint } from "./embeddedFontFingerprint";

test("font metadata does not multiply paragraph cache key storage", () => {
  const fonts = Array.from(
    { length: 25_000 },
    (_, index) =>
      ({
        name: `Face-${index}-${"a".repeat(100)}`,
        family: "roman",
        pitch: "variable",
        charset: { val: "00", characterSet: "ANSI" },
      }) satisfies FontInfo,
  );
  const documentFontFacts = buildDocumentFontFacts({ fonts });
  if (!documentFontFacts)
    throw new Error("Expected populated document font facts");
  expect(documentFontFacts.cacheIdentity.length).toBeLessThanOrEqual(12);

  const keys = Array.from({ length: 64 }, (_, index) => {
    const paragraph = {
      kind: "paragraph",
      id: String(index),
      attrs: { documentFontFacts },
      runs: [{ kind: "text", text: `Paragraph ${index}`, documentFontFacts }],
    } as const satisfies ParagraphBlock;
    return hashParagraphBlock(paragraph);
  });
  expect(keys.reduce((total, key) => total + key.length * 2, 0)).toBeLessThan(
    128 * 1024,
  );
  expect(keys.every((key) => !key.includes("Face-0-"))).toBe(true);
});

test("embedded font metadata also uses a bounded cache identity", () => {
  const families = new Map(
    Array.from(
      { length: 25_000 },
      (_, index) =>
        [`Font-${index}`, `Embedded-${index}-${"a".repeat(100)}`] as const,
    ),
  );
  expect(embeddedFontFingerprint(families).length).toBeLessThanOrEqual(12);
});
