import { expect, test } from "bun:test";

import { normalizeRevisionIdsInXmlParts, REVISION_ELEMENT_NAMES } from "./revisionIdNormalization";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

test("keeps unique revision ids and paired range ids byte-identical", () => {
  const xml = `<x:document xmlns:x="${W}"><x:moveFromRangeStart x:id="4" x:name="m"/><x:moveFrom x:id="7" x:author="A"/><x:moveFromRangeEnd x:id="4"/></x:document>`;
  expect(
    normalizeRevisionIdsInXmlParts(new Map([["word/document.xml", xml]])).get("word/document.xml"),
  ).toBe(xml);
});

test("assigns fresh ids to repeated physical revisions across parts idempotently", () => {
  const parts = new Map([
    [
      "word/document.xml",
      `<?xml version="1.0"?>\r\n<x:document xmlns:x="${W}"><?keep value?><x:del x:id="9" x:author="A &amp; &quot;B&quot;"/><!-- exact --><![CDATA[<x:del x:id="9"/>]]><x:del x:id='9'/><x:bookmarkStart x:id="9"/></x:document>`,
    ],
    ["word/header1.xml", `<ž:hdr xmlns:ž="${W}"><ž:ins ž:id="9"/></ž:hdr>`],
  ]);
  const once = normalizeRevisionIdsInXmlParts(parts);
  const twice = normalizeRevisionIdsInXmlParts(once);
  expect(once.get("word/document.xml")).toContain("<x:del x:id='0'/>");
  expect(once.get("word/header1.xml")).toContain('<ž:ins ž:id="1"/>');
  expect(twice).toEqual(once);
  expect(once.get("word/document.xml")).toContain(
    '<?keep value?><x:del x:id="9" x:author="A &amp; &quot;B&quot;"/><!-- exact --><![CDATA[<x:del x:id="9"/>]]>',
  );
  expect(once.get("word/document.xml")).toContain('<x:bookmarkStart x:id="9"/>');
});

test("normalizes every tracked revision element kind", () => {
  const revisions = [...REVISION_ELEMENT_NAMES].map((name) => `<w:${name} w:id="5"/>`).join("");
  const xml = `<w:document xmlns:w="${W}">${revisions}</w:document>`;
  const normalized = normalizeRevisionIdsInXmlParts(new Map([["word/document.xml", xml]])).get(
    "word/document.xml",
  );
  const ids = [...(normalized ?? "").matchAll(/w:id="(\d+)"/gu)].map((match) => match[1]);
  expect(ids).toHaveLength(REVISION_ELEMENT_NAMES.size);
  expect(new Set(ids).size).toBe(ids.length);
  expect(ids.at(0)).toBe("5");
});
