/**
 * An empty run property set is not an absent one.
 *
 * `w:rPr` is optional on every one of its owners, so a producer that wrote
 * `<w:rPr/>` stated something an absent element does not. The reader keyed the
 * record on the properties the element yielded and answered `undefined` for one
 * that yielded none, so the empty element reached no model and no writer put it
 * back. That is the rule `w:tblPrEx`, `w:trPr` and `w:tcPr` already follow:
 * presence is the value.
 *
 * It cannot state formatting — that is the whole point — but on the paragraph
 * mark it is also where `w:rPrChange`, `w:ins`, `w:del`, `w:moveFrom` and
 * `w:moveTo` live, so an empty one is pure presence and nothing else.
 *
 * Every owner is exercised, because the point of one reader and one writer is
 * that they cannot answer differently: a run, the paragraph mark, a style, a
 * numbering level and the snapshot inside either kind of `w:rPrChange`. The two
 * snapshots came back before this change as well — `CT_RPrChange` requires a
 * `w:rPr` and the run writer repairs a missing one, and the paragraph mark's
 * revision is a capture — and they are here so the family cannot drift apart
 * again. `w:rPrDefault` is the same repaired case, in `stylesSerializer`.
 *
 * The editor leg is here too: the paragraph mark's record rides
 * `ParagraphAttrs._originalFormatting.runProperties`, and a projection that
 * collapsed `{}` to absent would lose on the way back what the save keeps.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import type { Document, Paragraph } from "../types/document";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createEmptyDocument } from "../utils/createDocument";
import { parseNumbering } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import { parseRun } from "./runParser";
import { serializeNumberingXml } from "./serializer/numberingSerializer";
import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { serializeRun } from "./serializer/runSerializer";
import { serializeStylesXml } from "./serializer/stylesSerializer";
import { serializeTextFormatting } from "./serializer/textFormattingSerializer";
import { parseStyleDefinitions } from "./styleParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const parseOne = (xml: string): XmlElement => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("the fixture did not parse");
  }
  return root;
};

/**
 * One owner of a run property set.
 *
 * `save` takes the owner's markup and writes it back from the model alone,
 * which is the same forcing the survival law's L2 applies: folio replays
 * captured bytes whenever a fingerprint says the model still agrees with them,
 * so a round trip that went through the capture would exercise the replay
 * rather than the writer and prove nothing about either.
 *
 * `authored` is the owner holding an empty `w:rPr`, `expected` is the fragment
 * that has to come back, and `reauthor` puts what the save wrote back into the
 * shape `save` reads, for the second save.
 */
type RunPropertySetOwner = Readonly<{
  name: string;
  authored: string;
  expected: string;
  save: (xml: string) => string;
  reauthor: (saved: string) => string;
}>;

const saveRun = (xml: string): string => serializeRun(parseRun(parseOne(xml), null, null));

const saveParagraphMark = (xml: string): string => {
  const paragraph = parseParagraph(parseOne(xml), null, null, null);
  return serializeParagraphFormatting(
    paragraph.formatting,
    paragraph.propertyChanges,
    paragraph.pPrMark,
  );
};

const documentHolding = (paragraph: Paragraph): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph];
  return document;
};

/** The mark's properties again, after the editor has had the paragraph. */
const saveParagraphMarkThroughEditor = (xml: string): string => {
  const document = documentHolding(parseParagraph(parseOne(xml), null, null, null));
  const projected = fromProseDoc(toProseDoc(document), document);
  const roundTripped = projected.package.document.content.at(0);
  if (roundTripped?.type !== "paragraph") {
    throw new Error("the projection did not hand back a paragraph");
  }
  return serializeParagraphFormatting(
    roundTripped.formatting,
    roundTripped.propertyChanges,
    roundTripped.pPrMark,
  );
};

const saveStyles = (xml: string): string => serializeStylesXml(parseStyleDefinitions(xml, null));

const saveNumbering = (xml: string): string =>
  serializeNumberingXml(parseNumbering(xml).definitions);

const wrapRun = (inner: string): string => `<w:r xmlns:w="${W}">${inner}<w:t>x</w:t></w:r>`;
const wrapParagraph = (inner: string): string => `<w:p xmlns:w="${W}">${inner}</w:p>`;

const REVISION = 'w:id="7" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"';

/** A run writer emits `<w:r>` with no bindings; the reader needs them back. */
const rebindRun = (saved: string): string => saved.replace("<w:r>", `<w:r xmlns:w="${W}">`);

/** A part writer emits the whole part, bindings included. */
const asWritten = (saved: string): string => saved;

const OWNERS = [
  {
    name: "a run",
    authored: wrapRun("<w:rPr/>"),
    expected: "<w:rPr/>",
    save: saveRun,
    reauthor: rebindRun,
  },
  {
    name: "the paragraph mark",
    authored: wrapParagraph("<w:pPr><w:rPr/></w:pPr>"),
    expected: "<w:rPr/>",
    save: saveParagraphMark,
    reauthor: wrapParagraph,
  },
  {
    name: "the paragraph mark through the editor",
    authored: wrapParagraph("<w:pPr><w:rPr/></w:pPr>"),
    expected: "<w:rPr/>",
    save: saveParagraphMarkThroughEditor,
    reauthor: wrapParagraph,
  },
  {
    name: "a style",
    authored:
      `<w:styles xmlns:w="${W}">` +
      '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:rPr/></w:style>' +
      "</w:styles>",
    expected: '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:rPr/>',
    save: saveStyles,
    reauthor: asWritten,
  },
  {
    name: "document defaults",
    authored:
      `<w:styles xmlns:w="${W}">` +
      "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>" +
      "</w:styles>",
    expected: "<w:rPrDefault><w:rPr/></w:rPrDefault>",
    save: saveStyles,
    reauthor: asWritten,
  },
  {
    name: "a numbering level",
    authored:
      `<w:numbering xmlns:w="${W}">` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
      '<w:numFmt w:val="bullet"/><w:lvlText w:val="&#183;"/><w:rPr/>' +
      "</w:lvl></w:abstractNum>" +
      "</w:numbering>",
    expected: "<w:rPr/>",
    save: saveNumbering,
    reauthor: asWritten,
  },
  {
    name: "the snapshot inside a run's tracked property change",
    authored: wrapRun(`<w:rPr><w:rPrChange ${REVISION}><w:rPr/></w:rPrChange></w:rPr>`),
    expected: `<w:rPrChange ${REVISION}><w:rPr/></w:rPrChange>`,
    save: saveRun,
    reauthor: rebindRun,
  },
  {
    name: "the snapshot inside the paragraph mark's tracked property change",
    authored: wrapParagraph(
      `<w:pPr><w:rPr><w:rPrChange ${REVISION}><w:rPr/></w:rPrChange></w:rPr></w:pPr>`,
    ),
    expected: "<w:rPr/></w:rPrChange>",
    save: saveParagraphMark,
    reauthor: wrapParagraph,
  },
] as const satisfies readonly RunPropertySetOwner[];

describe("an empty run property set is not an absent one", () => {
  test("every owner writes back the empty element it was given", () => {
    fc.assert(
      fc.property(fc.constantFrom(...OWNERS), (owner) => {
        expect(owner.save(owner.authored)).toContain(owner.expected);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("the second save is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...OWNERS), (owner) => {
        const once = owner.save(owner.authored);
        expect(owner.save(owner.reauthor(once))).toBe(once);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("an owner that never carried a run property set still writes none", () => {
    expect(saveRun(wrapRun(""))).not.toContain("<w:rPr");
    expect(saveParagraphMark(wrapParagraph("<w:pPr/>"))).not.toContain("<w:rPr");
    expect(saveParagraphMark(wrapParagraph(""))).toBe("");
    expect(saveParagraphMarkThroughEditor(wrapParagraph("<w:pPr/>"))).not.toContain("<w:rPr");
  });

  test("a constructed record whose fields emit nothing does not invent an empty set", () => {
    expect(
      serializeTextFormatting({
        color: {},
        fontFamily: {},
        language: {},
        styleId: "",
      }),
    ).toBe("");
  });

  test("the paragraph mark's empty record reaches the editor's attr and comes back", () => {
    const paragraph = parseParagraph(
      parseOne(wrapParagraph("<w:pPr><w:rPr/></w:pPr>")),
      null,
      null,
      null,
    );
    expect(paragraph.formatting?.runProperties).toEqual({});

    const node = toProseDoc(documentHolding(paragraph)).firstChild;
    // The attr has to say "present and empty": an `undefined` here is the
    // projection collapsing presence to absence, and the save after it writes
    // nothing however well the reader did.
    expect(node?.attrs["_originalFormatting"]).toEqual({ runProperties: {} });
  });
});
