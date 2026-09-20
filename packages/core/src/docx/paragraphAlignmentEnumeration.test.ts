/**
 * Every `ST_Jc` member survives a parse, a save and the editor projection.
 *
 * `narrowEnum` returns `undefined` for a member the picklist omits and the
 * caller drops the attribute, so an alignment the model does not spell is lost
 * on the next save. `PARAGRAPH_ALIGNMENT_VALUES` used to spell nine of the
 * twelve: `start`, `end` and `numTab` were dropped, and `<w:jc w:val="start"/>`
 * — what a Strict-profile producer writes for the direction-aware left — came
 * back unaligned.
 *
 * The sweep is over the generated list rather than a list spelled here, so a
 * schema refresh that adds a member widens the test with it.
 */

import { describe, expect, test } from "bun:test";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, ParagraphAlignment } from "../types/document";
import { PARAGRAPH_ALIGNMENT_VALUES } from "../types/documentEnumValues";

import { parseParagraphProperties } from "./paragraphParser";
import { parseStyles } from "./styleParser";
import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const parseParagraphTier = (alignment: string): ParagraphAlignment | undefined => {
  const pPr = parseXmlDocument(`<w:pPr ${WORD_NAMESPACE}><w:jc w:val="${alignment}"/></w:pPr>`);
  if (!pPr) {
    throw new Error("fixture did not parse");
  }
  return parseParagraphProperties(pPr, null)?.alignment;
};

const parseStyleTier = (alignment: string): ParagraphAlignment | undefined =>
  parseStyles(
    `<w:styles ${WORD_NAMESPACE}>
       <w:style w:type="paragraph" w:styleId="Aligned">
         <w:name w:val="Aligned"/>
         <w:pPr><w:jc w:val="${alignment}"/></w:pPr>
       </w:style>
     </w:styles>`,
    null,
  ).get("Aligned")?.pPr?.alignment;

const documentWithAlignment = (alignment: ParagraphAlignment): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          formatting: { alignment },
          content: [{ type: "run", content: [{ type: "text", text: "Aligned" }] }],
        },
      ],
    },
  },
});

const firstParagraphAlignment = (document: Document): ParagraphAlignment | undefined => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }
  return block.formatting?.alignment;
};

describe("ST_Jc", () => {
  // The sweeps below run over the model's own list, so they shrink with it.
  // `scripts/narrowed-enum-schema-types.test.ts` is what holds that list to the
  // enumeration; this names the three members it used to be missing.
  test("start, end and numTab are members", () => {
    expect(PARAGRAPH_ALIGNMENT_VALUES).toContain("start");
    expect(PARAGRAPH_ALIGNMENT_VALUES).toContain("end");
    expect(PARAGRAPH_ALIGNMENT_VALUES).toContain("numTab");
  });

  test.each(PARAGRAPH_ALIGNMENT_VALUES)("a paragraph's w:jc reads %s", (alignment) => {
    expect(parseParagraphTier(alignment)).toBe(alignment);
  });

  test.each(PARAGRAPH_ALIGNMENT_VALUES)("a style's w:jc reads %s", (alignment) => {
    expect(parseStyleTier(alignment)).toBe(alignment);
  });

  test.each(PARAGRAPH_ALIGNMENT_VALUES)("%s survives parse, save and parse", (alignment) => {
    const saved = serializeParagraphFormatting({ alignment });
    expect(saved).toContain(`<w:jc w:val="${alignment}"/>`);

    const reparsed = parseXmlDocument(saved.replace("<w:pPr>", `<w:pPr ${WORD_NAMESPACE}>`));
    if (!reparsed) {
      throw new Error("the saved properties did not parse");
    }
    expect(parseParagraphProperties(reparsed, null)?.alignment).toBe(alignment);
  });

  test.each(PARAGRAPH_ALIGNMENT_VALUES)("%s survives the editor projection", (alignment) => {
    const original = documentWithAlignment(alignment);
    expect(firstParagraphAlignment(fromProseDoc(toProseDoc(original), original))).toBe(alignment);
  });

  // `start` is not `left`: a reader that aliased the two would pass every
  // assertion above and still write the wrong alignment into an RTL paragraph.
  test("start and end are members, not spellings of left and right", () => {
    expect(parseParagraphTier("start")).toBe("start");
    expect(parseParagraphTier("end")).toBe("end");
    expect(serializeParagraphFormatting({ alignment: "start" })).toContain('<w:jc w:val="start"/>');
  });

  test("a token outside the enumeration does not become an alignment", () => {
    expect(parseParagraphTier("sideways")).toBeUndefined();
  });
});
