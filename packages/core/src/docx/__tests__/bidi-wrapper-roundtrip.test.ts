/**
 * `w:bdo` and `w:dir` are the Unicode bidirectional controls, and folio
 * discarded both.
 *
 * `w:dir` is an embedding: the text inside is laid out in the given direction
 * and the bidirectional algorithm still resolves the characters within it.
 * `w:bdo` is an override: the algorithm is switched off and every character is
 * laid out in the given direction, which is what makes a Latin word inside an
 * `rtl` override read backwards. Dropping either changes what the reader sees,
 * not only what the file says, and in a right-to-left document that is the
 * difference between a readable line and a scrambled one.
 *
 * They are transparent inline wrappers — they constrain how their content is
 * laid out and nothing else — so they nest, they hold anything paragraph
 * content holds, and everything that walks a paragraph reads straight through
 * them.
 */

import { describe, expect, test } from "bun:test";

import { getParagraphText, parseParagraph } from "../paragraphParser";
import { serializeParagraph } from "../serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "../xmlParser";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

const paragraphFrom = (inner: string) => {
  const node = parseXmlDocument(`<w:p ${NS}>${inner}</w:p>`) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return parseParagraph(node, new Map(), null, null);
};

const roundTrip = (inner: string): string => serializeParagraph(paragraphFrom(inner));

describe("a bidirectional wrapper survives a rebuild", () => {
  test("an override keeps its direction and its runs", () => {
    const xml = roundTrip(
      '<w:r><w:t>abc</w:t></w:r><w:bdo w:val="rtl"><w:r><w:t>def</w:t></w:r></w:bdo>',
    );
    expect(xml).toContain('<w:bdo w:val="rtl"><w:r><w:t>def</w:t></w:r></w:bdo>');
  });

  test("an embedding is written as w:dir, not as an override", () => {
    const xml = roundTrip('<w:dir w:val="ltr"><w:r><w:t>abc</w:t></w:r></w:dir>');
    expect(xml).toContain('<w:dir w:val="ltr">');
    expect(xml).not.toContain("<w:bdo");
  });

  test("a wrapper that states no direction keeps stating none", () => {
    const xml = roundTrip("<w:bdo><w:r><w:t>abc</w:t></w:r></w:bdo>");
    expect(xml).toContain("<w:bdo><w:r><w:t>abc</w:t></w:r></w:bdo>");
  });

  test("nesting survives, inside and out", () => {
    const xml = roundTrip(
      '<w:dir w:val="rtl"><w:r><w:t>a</w:t></w:r><w:bdo w:val="ltr"><w:r><w:t>b</w:t></w:r></w:bdo></w:dir>',
    );
    expect(xml).toContain(
      '<w:dir w:val="rtl"><w:r><w:t>a</w:t></w:r><w:bdo w:val="ltr"><w:r><w:t>b</w:t></w:r></w:bdo></w:dir>',
    );
  });

  test("a non-run child keeps its place and its order", () => {
    const xml = roundTrip(
      '<w:bdo w:val="rtl"><w:bookmarkStart w:id="1" w:name="bm"/><m:oMath/><w:r><w:t>z</w:t></w:r></w:bdo>',
    );
    expect(xml).toContain(
      '<w:bookmarkStart w:id="1" w:name="bm"/><m:oMath/><w:r><w:t>z</w:t></w:r>',
    );
  });

  test("plain-text extraction reads through the wrapper", () => {
    const paragraph = paragraphFrom(
      '<w:r><w:t>abc</w:t></w:r><w:bdo w:val="rtl"><w:r><w:t>def</w:t></w:r></w:bdo>',
    );
    expect(getParagraphText(paragraph)).toBe("abcdef");
  });
});

/**
 * A wrapper and a revision nest either way round, and the author picked which.
 *
 * Nothing about `w:bdo` inside `w:ins` is less legal than `w:ins` inside
 * `w:bdo`; they say different things, and a save that turns one into the other
 * moves text into or out of the revision. Serializing the parsed paragraph is
 * the forced path for inline content, the same one the corpus reserialize
 * invariant forces for captured slots: folio replays a captured `w:pPr` and a
 * captured `w:sdtPr`, never a run or a wrapper, so what comes back here is
 * what the serializer built.
 */
describe("a revision and a bidirectional wrapper keep the order they were authored in", () => {
  const ATTRS = 'w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"';

  const AUTHORED_ORDERS = [
    `<w:ins ${ATTRS}><w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo></w:ins>`,
    `<w:bdo w:val="rtl"><w:ins ${ATTRS}><w:r><w:t>x</w:t></w:r></w:ins></w:bdo>`,
    `<w:del ${ATTRS}><w:dir w:val="ltr"><w:r><w:delText>x</w:delText></w:r></w:dir></w:del>`,
    `<w:dir w:val="ltr"><w:del ${ATTRS}><w:r><w:delText>x</w:delText></w:r></w:del></w:dir>`,
  ];

  for (const authored of AUTHORED_ORDERS) {
    test(`${authored} comes back unchanged`, () => {
      expect(roundTrip(authored)).toBe(`<w:p>${authored}</w:p>`);
    });

    test(`${authored} is a fixed point of a second save`, () => {
      const once = roundTrip(authored);
      const twice = serializeParagraph(
        parseParagraph(
          parseXmlDocument(once.replace("<w:p>", `<w:p ${NS}>`)) as XmlElement,
          new Map(),
          null,
          null,
        ),
      );
      expect(twice).toBe(once);
    });
  }
});
