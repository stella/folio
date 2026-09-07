import { expect, test } from "bun:test";

import {
  normalizeRevisionIdsInXmlParts,
  RevisionIdCollisionError,
  REVISION_ELEMENT_NAMES,
} from "./revisionIdNormalization";

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

test("respects Strict namespaces and nested prefix rebinding", () => {
  const strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
  const xml = `<w:document xmlns:w="${strict}"><w:del w:id="8"/><w:del xmlns:w="urn:foreign" w:id="8"/><w:del w:id="8"/></w:document>`;
  const normalized = normalizeRevisionIdsInXmlParts(new Map([["word/document.xml", xml]]));
  expect(normalized.get("word/document.xml")).toBe(
    xml.replace('<w:del w:id="8"/></w:document>', '<w:del w:id="0"/></w:document>'),
  );
});

// Every id the pass lets stand or mints goes through one choke point, which
// throws rather than emit a package where two revisions answer to one `w:id`.
test("claims every id once when every part repeats the same one", () => {
  const parts = new Map(
    ["word/document.xml", "word/header1.xml", "word/footnotes.xml"].map((path) => [
      path,
      `<w:root xmlns:w="${W}"><w:del w:id="3"/><w:ins w:id="3"/><w:rPrChange w:id="3"/></w:root>`,
    ]),
  );

  expect(() => normalizeRevisionIdsInXmlParts(parts)).not.toThrow(RevisionIdCollisionError);

  const ids = [...normalizeRevisionIdsInXmlParts(parts).values()].flatMap((xml) =>
    [...xml.matchAll(/w:id="(\d+)"/gu)].map((match) => match[1]),
  );
  expect(ids).toHaveLength(9);
  expect(new Set(ids).size).toBe(ids.length);
});
