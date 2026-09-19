/**
 * A paraId is identified by its namespace URI, at every reader of the key.
 *
 * Three Word generations write the key (`w14`, `w15`, `w16cex`) and a producer
 * may bind any of them to a prefix of its own. Reading by prefix misses those
 * files; reading by local name alone picks up an unrelated `vendor:paraId` and
 * threads the comment into a stranger's conversation.
 */

import { describe, expect, test } from "bun:test";

import { parseComments, parseCommentsExtended } from "./commentParser";
import { parseParagraph } from "./paragraphParser";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CEX_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const FOREIGN_NS = "https://vendor.example/ns";

const PARA_ID = "1A2B3C4D";

const paraIdOf = (declaration: string, attribute: string): string | undefined => {
  const root = parseXmlDocument(
    `<w:p xmlns:w="${W_NS}" ${declaration} ${attribute}><w:r><w:t>body</w:t></w:r></w:p>`,
  ) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse paragraph XML fixture");
  }
  return parseParagraph(root, null, null, null, null, null).paraId;
};

describe("a paragraph's paraId", () => {
  for (const [generation, uri] of [
    ["Word 2010 (w14)", W14_NS],
    ["Word 2012 (w15)", W15_NS],
    ["Word 2018 (w16cex)", W16CEX_NS],
  ] as const) {
    test(`is read under any prefix bound to ${generation}`, () => {
      expect(paraIdOf(`xmlns:ns0="${uri}"`, `ns0:paraId="${PARA_ID}"`)).toBe(PARA_ID);
    });
  }

  test("is not read from a foreign namespace", () => {
    expect(paraIdOf(`xmlns:vendor="${FOREIGN_NS}"`, `vendor:paraId="${PARA_ID}"`)).toBeUndefined();
  });

  test("is still read from an undeclared conventional prefix", () => {
    // Malformed, and Word reads it: no URI lookup can resolve an unbound prefix.
    expect(paraIdOf("", `w14:paraId="${PARA_ID}"`)).toBe(PARA_ID);
  });
});

describe("the commentsExtensible join key", () => {
  const commentsXml = `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:comment w:id="1" w:author="Alice" w:date="2024-02-10T15:30:00" w14:paraId="${PARA_ID}">
    <w:p w14:paraId="${PARA_ID}"><w:r><w:t>First comment</w:t></w:r></w:p>
  </w:comment>
</w:comments>`;

  const dateFor = (extensibleXml: string): string | undefined =>
    parseComments(commentsXml, null, null, {}, new Map(), extensibleXml).at(0)?.date;

  test("is read under any prefix bound to the 2018 namespace", () => {
    expect(
      dateFor(`<ns0:commentsExtensible xmlns:ns0="${W16CEX_NS}">
  <ns0:comment ns0:paraId="${PARA_ID}" ns0:dateUtc="2024-02-10T14:30:00Z"/>
</ns0:commentsExtensible>`),
    ).toBe("2024-02-10T14:30:00Z");
  });

  test("is not read from a foreign namespace", () => {
    // The local date stands: a stranger's key must not carry a UTC timestamp in.
    expect(
      dateFor(`<w16cex:commentsExtensible xmlns:w16cex="${W16CEX_NS}" xmlns:vendor="${FOREIGN_NS}">
  <w16cex:comment vendor:paraId="${PARA_ID}" vendor:dateUtc="2024-02-10T14:30:00Z"/>
</w16cex:commentsExtensible>`),
    ).toBe("2024-02-10T15:30:00");
  });
});

describe("the commentsExtended thread link", () => {
  const PARENT_PARA_ID = "5E6F7A8B";

  test("is read under any prefix bound to the 2012 namespace", () => {
    const info = parseCommentsExtended(`<ns0:commentsEx xmlns:ns0="${W15_NS}">
  <ns0:commentEx ns0:paraId="${PARA_ID}" ns0:paraIdParent="${PARENT_PARA_ID}" ns0:done="1"/>
</ns0:commentsEx>`);
    expect(info.get(PARA_ID)).toEqual({ parentParaId: PARENT_PARA_ID, done: true });
  });

  test("is not read from a foreign namespace", () => {
    const info =
      parseCommentsExtended(`<w15:commentsEx xmlns:w15="${W15_NS}" xmlns:vendor="${FOREIGN_NS}">
  <w15:commentEx w15:paraId="${PARA_ID}" vendor:paraIdParent="${PARENT_PARA_ID}"/>
</w15:commentsEx>`);
    expect(info.get(PARA_ID)?.parentParaId).toBeUndefined();
  });
});
