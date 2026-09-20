/**
 * A revision keeps the wrapper it was authored around.
 *
 * `w:bdo`, `w:dir` and an inline `w:sdt` are transparent: they state how
 * their content is laid out or what it is bound to, and nothing about the
 * revision. A parser that cannot hold one inside a revision lifts it out to a
 * sibling, and the wrapped text goes with it: `<w:ins><w:bdo>x</w:bdo></w:ins>`
 * saved as `<w:ins/><w:bdo>x</w:bdo>` says `x` was never inserted, so
 * accepting the revision keeps `x` and rejecting it keeps `x` too.
 *
 * Both authored orders are legal and both have to come back unchanged: the
 * revision may hold the wrapper, and the wrapper may hold the revision.
 */

import { describe, expect, test } from "bun:test";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

const roundTrip = (inner: string): string => {
  const node = parseXmlDocument(`<w:p ${NS}>${inner}</w:p>`) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return serializeParagraph(parseParagraph(node, new Map(), null, null));
};

const REVISIONS = [
  { tag: "ins", removes: false },
  { tag: "del", removes: true },
  { tag: "moveFrom", removes: true },
  { tag: "moveTo", removes: false },
] as const;

const WRAPPERS = [
  { name: "w:bdo", open: '<w:bdo w:val="rtl">', close: "</w:bdo>" },
  { name: "w:dir", open: '<w:dir w:val="rtl">', close: "</w:dir>" },
  {
    name: "inline w:sdt",
    open: "<w:sdt><w:sdtPr/><w:sdtContent>",
    close: "</w:sdtContent></w:sdt>",
  },
] as const;

const ATTRS = 'w:id="7" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';

describe("a revision keeps the wrapper it was authored around", () => {
  for (const { tag, removes } of REVISIONS) {
    for (const { name, open, close } of WRAPPERS) {
      const text = removes ? "<w:delText>x</w:delText>" : "<w:t>x</w:t>";

      test(`${name} stays inside w:${tag}`, () => {
        const xml = roundTrip(
          `<w:${tag} ${ATTRS}>${open}<w:r><w:t>x</w:t></w:r>${close}</w:${tag}>`,
        );
        expect(xml).toContain(`<w:${tag} ${ATTRS}>${open}<w:r>${text}</w:r>${close}</w:${tag}>`);
      });

      test(`w:${tag} stays inside ${name}`, () => {
        const xml = roundTrip(
          `${open}<w:${tag} ${ATTRS}><w:r><w:t>x</w:t></w:r></w:${tag}>${close}`,
        );
        expect(xml).toContain(`${open}<w:${tag} ${ATTRS}><w:r>${text}</w:r></w:${tag}>${close}`);
      });
    }
  }

  test("a revision that removes content writes delText through the wrapper", () => {
    const xml = roundTrip(
      `<w:del ${ATTRS}><w:bdo w:val="rtl"><w:r><w:t>gone</w:t></w:r></w:bdo></w:del>`,
    );
    expect(xml).toContain("<w:delText>gone</w:delText>");
    expect(xml).not.toContain("<w:t>gone</w:t>");
  });

  test("a wrapper nested three deep keeps every level inside the revision", () => {
    const inner =
      `<w:ins ${ATTRS}><w:bdo w:val="rtl"><w:sdt><w:sdtPr/><w:sdtContent>` +
      '<w:dir w:val="ltr"><w:r><w:t>x</w:t></w:r></w:dir>' +
      "</w:sdtContent></w:sdt></w:bdo></w:ins>";
    expect(roundTrip(inner)).toContain(inner);
  });

  test("the control inside a revision is rebuilt, not replayed", () => {
    // The inline control is the one wrapper here with a captured slot
    // (`rawPropertiesXml`). Clearing it is what the corpus reserialize
    // invariant does to every rebuildable capture: what comes back then is
    // the serializer's work rather than the parser's bytes.
    const node = parseXmlDocument(
      `<w:p ${NS}><w:ins ${ATTRS}><w:sdt><w:sdtPr><w:tag w:val="bound"/></w:sdtPr>` +
        "<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:ins></w:p>",
    ) as XmlElement | null;
    if (!node) {
      throw new Error("the fixture did not parse");
    }
    const paragraph = parseParagraph(node, new Map(), null, null);
    const revision = paragraph.content.at(0);
    const sdt = revision?.type === "insertion" ? revision.content.at(0) : undefined;
    if (sdt?.type !== "inlineSdt") {
      throw new Error("the control was not kept inside the revision");
    }
    sdt.properties.rawPropertiesXml = undefined;

    const xml = serializeParagraph(paragraph);
    expect(xml).toContain(`<w:ins ${ATTRS}><w:sdt>`);
    expect(xml).toContain("<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:ins>");
  });

  test("content the revision may not hold still splits it", () => {
    const xml = roundTrip(
      `<w:ins ${ATTRS}><w:r><w:t>a</w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>b</w:t></w:r></w:ins>`,
    );
    expect(xml).toContain(
      `<w:ins ${ATTRS}><w:r><w:t>a</w:t></w:r></w:ins><w:commentRangeStart w:id="1"/><w:ins ${ATTRS}><w:r><w:t>b</w:t></w:r></w:ins>`,
    );
  });
});
