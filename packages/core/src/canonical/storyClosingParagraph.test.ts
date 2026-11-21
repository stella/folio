/**
 * Which stories the session closes with a paragraph: one whose last
 * paragraph or table, read through content controls and custom XML, is a
 * table, or one with neither. A control or wrapper holding neither, at any
 * depth, is passed over, so the block before it decides.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document } from "../types/document";
import { normalizeCanonicalSource } from "./canonicalSourceNormalization";

const paragraph = (paraId: string): BlockContent => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: paraId }] }],
});
const control = (content: BlockContent[]): BlockContent => ({
  type: "blockSdt",
  properties: { sdtType: "richText" },
  content,
});
const table: BlockContent = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph("20000001")] }],
    },
  ],
};

const closes = (content: BlockContent[]): boolean => {
  const document: Document = { package: { document: { content } } };
  return normalizeCanonicalSource(document).normalized.some(
    ({ story, minted }) => story === "main" && minted.length > 0,
  );
};

test.each([
  ["a paragraph", [paragraph("10000001")], false],
  [
    "a paragraph, then an empty control",
    [paragraph("10000001"), control([])],
    false,
  ],
  [
    "a paragraph, then a control holding an empty control",
    [paragraph("10000001"), control([control([])])],
    false,
  ],
  ["a table", [paragraph("10000001"), table], true],
  [
    "a control holding a table",
    [paragraph("10000001"), control([control([table])])],
    true,
  ],
  ["a table, then an empty control", [table, control([control([])])], true],
  ["nothing but an empty control", [control([])], true],
] as const)(
  "a story ending in %s is closed: %p",
  (_name, content, expected) => {
    expect(closes([...content])).toBe(expected);
  },
);
