/**
 * A link and a simple field keep the children folio does not model.
 *
 * `CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`, so either may
 * hold a permission range, a proofing error, a transparent wrapper or one of
 * the eight custom-XML revision ranges between its runs. folio modelled the
 * run and the bookmark boundaries and let the other twenty-nine names fall off
 * the end of a `switch`, so a link around a clause with a protection range in
 * it came back without the range.
 *
 * Position matters for the same reason it does inside a tracked change: a
 * capture written beside the link rather than inside it survives the save and
 * still moves the markup out of the range the author drew, so every assertion
 * here is about the markup being *inside* the element it was read from.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The two `EG_PContent` containers with a parser of their own. */
const CONTAINERS = [
  { tag: "hyperlink", attributes: ' w:anchor="clause"' },
  { tag: "fldSimple", attributes: ' w:instr="PAGE"' },
] as const;

/**
 * Children either content model declares and folio models nothing of.
 *
 * One per shape rather than all of them: an empty marker, a marker carrying
 * revision attributes, a container with text inside it, and an element from a
 * namespace the content model does not name at all, which is the one the
 * dispatcher's sink rather than its handler map has to catch.
 *
 * The four transparent wrappers are not here: both containers model them now,
 * so `inlineWrapperInsideLinkAndField.property.test.ts` asserts about them
 * instead, including the canonical re-nesting the editor leg performs.
 */
const UNMODELLED_CHILDREN = [
  '<w:permStart w:id="7" w:edGrp="everyone"/>',
  '<w:proofErr w:type="spellStart"/>',
  '<w:customXmlInsRangeStart w:id="3" w:author="Reviewer"/>',
  "<w:sdt><w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent></w:sdt>",
  '<x:note xmlns:x="urn:example:vendor" x:kind="aside">kept</x:note>',
] as const;

const parseParagraphXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse the paragraph fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
};

const paragraphXml = (container: (typeof CONTAINERS)[number], child: string): string =>
  `<w:p xmlns:w="${W}">` +
  `<w:${container.tag}${container.attributes}>` +
  `<w:r><w:t>text</w:t></w:r>${child}` +
  `</w:${container.tag}></w:p>`;

/** What the container holds in the saved markup, or `null` when there is none. */
const inside = (savedXml: string, tag: string): string | null => {
  const open = savedXml.indexOf(`<w:${tag} `);
  const close = savedXml.indexOf(`</w:${tag}>`, open);
  if (open === -1 || close === -1) {
    return null;
  }
  return savedXml.slice(savedXml.indexOf(">", open) + 1, close);
};

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

/** The paragraph as it comes back from the editor, with nothing edited. */
const throughTheEditor = (paragraph: Paragraph): Paragraph => {
  const input = wrapParagraph(paragraph);
  const first = fromProseDoc(toProseDoc(input), input).package.document.content.at(0);
  if (first?.type !== "paragraph") {
    throw new Error("Expected the round trip to give a paragraph back");
  }
  return first;
};

describe("a link and a simple field keep the children folio does not model", () => {
  for (const container of CONTAINERS) {
    test(`w:${container.tag} keeps a w:permStart inside itself across a save`, () => {
      const paragraph = parseParagraphXml(paragraphXml(container, '<w:permStart w:id="7"/>'));

      const item = paragraph.content.at(0);
      const held = item?.type === "hyperlink" ? item.children : undefined;
      const fieldContent = item?.type === "simpleField" ? item.content : undefined;
      expect((held ?? fieldContent)?.some((child) => child.type === "preservedInline")).toBe(true);

      expect(inside(serializeParagraph(paragraph), container.tag)).toContain(
        '<w:permStart w:id="7"/>',
      );
    });
  }

  test("a child kept by one container is kept by both, for every kind", () => {
    fc.assert(
      fc.property(
        fc.constantFrom(...CONTAINERS),
        fc.constantFrom(...UNMODELLED_CHILDREN),
        (container, child) => {
          const saved = serializeParagraph(parseParagraphXml(paragraphXml(container, child)));
          expect(inside(saved, container.tag)).toContain(child);

          // Save, reopen, save: the second save is where a capture that only
          // replays and does not re-parse stops being a fixed point.
          const reparsed = serializeParagraph(
            parseParagraphXml(`<w:p xmlns:w="${W}">${saved.slice(saved.indexOf(">") + 1)}`),
          );
          expect(inside(reparsed, container.tag)).toContain(child);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the editor gives the child back inside the same container", () => {
    fc.assert(
      fc.property(
        fc.constantFrom(...CONTAINERS),
        fc.constantFrom(...UNMODELLED_CHILDREN),
        (container, child) => {
          const reopened = throughTheEditor(parseParagraphXml(paragraphXml(container, child)));
          expect(inside(serializeParagraph(reopened), container.tag)).toContain(child);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a link that also holds a revision keeps the capture in source order", () => {
    const paragraph = parseParagraphXml(
      `<w:p xmlns:w="${W}"><w:hyperlink w:anchor="clause">` +
        `<w:r><w:t>before</w:t></w:r>` +
        '<w:permStart w:id="7"/>' +
        '<w:ins w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">' +
        "<w:r><w:t>added</w:t></w:r></w:ins>" +
        "</w:hyperlink></w:p>",
    );

    // The revision is hoisted around the link and the capture is not, so the
    // capture has to land in the link that held it rather than at the end of
    // whatever segment the walk happened to be building.
    const saved = serializeParagraph(paragraph);
    expect(saved.indexOf("<w:permStart")).toBeLessThan(saved.indexOf("<w:ins "));
    expect(inside(saved, "hyperlink")).toContain('<w:permStart w:id="7"/>');
  });
});
