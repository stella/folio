import { expect, test } from "bun:test";

import { normalizeParaIdRangeInXmlParts, paraIdInRange } from "./paraIdRangeNormalization";

const PARA_ID = /\b(?:w|w14|w15|w16cid):(?:paraId|paraIdParent|textId)="([0-9A-Fa-f]{8})"/gu;

const paraIdsOf = (xml: string): string[] => [...xml.matchAll(PARA_ID)].map(([, id]) => id ?? "");

const inRange = (id: string): boolean => Number.parseInt(id, 16) < 0x8000_0000;

test("leaves a package whose ids already fit byte-identical", () => {
  const parts = new Map([
    ["word/document.xml", `<w:p w14:paraId="12AB34CD" w14:textId="12AB34CD"/>`],
  ]);
  expect(normalizeParaIdRangeInXmlParts(parts)).toEqual(parts);
});

test("brings an out-of-range paragraph id inside the 31-bit range", () => {
  const normalized = normalizeParaIdRangeInXmlParts(
    new Map([["word/document.xml", `<w:p w14:paraId="FFAABBCC" w14:textId="FFAABBCC"/>`]]),
  ).get("word/document.xml");

  const ids = paraIdsOf(normalized ?? "");
  expect(ids).toHaveLength(2);
  expect(ids.every(inRange)).toBe(true);
  // The text marker a paragraph carries beside its id stays equal to it.
  expect(new Set(ids).size).toBe(1);
});

test("rewrites a paragraph and every reference to it the same way", () => {
  const normalized = normalizeParaIdRangeInXmlParts(
    new Map([
      [
        "word/comments.xml",
        `<w:p w14:paraId="A0000001"/><w:p w14:paraId="FFFFFFF0"/><w:p w14:paraId="FFFFFFF1"/>`,
      ],
      [
        "word/commentsExtended.xml",
        `<w15:commentEx w15:paraId="FFFFFFF1" w15:paraIdParent="FFFFFFF0"/>`,
      ],
      ["word/commentsIds.xml", `<w16cid:commentId w16cid:paraId="FFFFFFF1"/>`],
    ]),
  );

  const comments = paraIdsOf(normalized.get("word/comments.xml") ?? "");
  const extended = paraIdsOf(normalized.get("word/commentsExtended.xml") ?? "");
  const durable = paraIdsOf(normalized.get("word/commentsIds.xml") ?? "");
  expect(comments.every(inRange)).toBe(true);
  // The reply still points at the comment it pointed at, and at its parent.
  expect(extended).toEqual([comments[2] ?? "", comments[1] ?? ""]);
  expect(durable).toEqual([comments[2] ?? ""]);
});

test("maps one value to one id wherever it is read", () => {
  expect(paraIdInRange("12AB34CD")).toBe("12AB34CD");
  expect(inRange(paraIdInRange("FFFFFFF0"))).toBe(true);
  expect(paraIdInRange("FFFFFFF0")).toBe(paraIdInRange("FFFFFFF0"));
  expect(paraIdInRange("FFFFFFF0")).not.toBe(paraIdInRange("FFFFFFF1"));
});

test("is idempotent", () => {
  const parts = new Map([
    ["word/document.xml", `<w:p w14:paraId="FFFFFFF0"/><w:p w14:paraId="F0000000"/>`],
  ]);
  const once = normalizeParaIdRangeInXmlParts(parts);
  expect(normalizeParaIdRangeInXmlParts(once)).toEqual(once);
});
