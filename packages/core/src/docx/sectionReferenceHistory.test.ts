import { expect, test } from "bun:test";
import type { SectionPropertyChange } from "../types/document";
import { parseSectionProperties } from "./sectionParser";
import { serializeSectionProperties } from "./serializer/sectionPropertiesSerializer";
import { parseXmlDocument } from "./xmlParser";
import { InvalidSectionReferenceHistoryError } from "./sectionReferenceHistory";

const parse = (xml: string) => parseSectionProperties(parseXmlDocument(
  `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${xml}</w:sectPr>`,
));
const revision = (previousReferences?: SectionPropertyChange["previousReferences"]): SectionPropertyChange => ({
  type: "sectionPropertyChange",
  info: {id: 7, author: "Reviewer", date: "2026-09-13T00:00:00.000Z"},
  previousProperties: { marginLeft: 1440 },
  ...(previousReferences !== undefined && {previousReferences}),
});

test("section history distinguishes native behavior from an explicitly empty prior selection", () => {
  for (const references of [undefined, {}, {
    headerReferences: [{type: "first" as const, rId: "rId8"}],
    footerReferences: [{type: "default" as const, rId: "rId9"}],
  }]) {
    const xml = serializeSectionProperties({propertyChanges: [revision(references)]});
    const parsed = parse(xml.slice("<w:sectPr>".length, -"</w:sectPr>".length));
    expect(parsed.propertyChanges?.at(0)?.previousReferences).toEqual(references);
  }
});

test("previous selections stay out of the native CT_SectPrBase payload", () => {
  const change = revision({ headerReferences: [{type: "default", rId: "rId1"}] });
  change.previousProperties = {headerReferences: [{type: "default", rId: "rId1"}]};
  const xml = serializeSectionProperties({propertyChanges: [change]});
  expect(xml).toContain("<w:sectPr/>");
  expect(xml.match(/<w:headerReference/gu)).toHaveLength(1);
  expect(xml).toContain('mc:Ignorable="frh"');
});

test("section history resolves an alternate namespace prefix", () => {
  const parsed = parse(`<w:sectPrChange w:id="7" w:author="Reviewer"><w:sectPr/><x:previousReferences xmlns:x="urn:stella:folio:section-reference-history:1"><w:footerReference w:type="even" r:id="rId2"/></x:previousReferences></w:sectPrChange>`);
  expect(parsed.propertyChanges?.at(0)?.previousReferences).toEqual({footerReferences: [{type: "even", rId: "rId2"}]});
});

test("malformed history refuses duplicate selections, invalid kinds, and missing relationships", () => {
  for (const children of [
    '<w:headerReference w:type="default" r:id="rId1"/><w:headerReference w:type="default" r:id="rId2"/>',
    '<w:headerReference w:type="unknown" r:id="rId1"/>',
    '<w:footerReference w:type="first"/>',
    '<w:unknown/>',
    '<w:headerReference w:type="default" xmlns:fake="urn:foreign" fake:id="rId1"/>',
    '<w:headerReference xmlns:fake="urn:foreign" fake:type="default" r:id="rId1"/>',
  ]) {
    expect(() => parse(`<w:sectPrChange w:id="7"><w:sectPr/><x:previousReferences xmlns:x="urn:stella:folio:section-reference-history:1">${children}</x:previousReferences></w:sectPrChange>`)).toThrow(InvalidSectionReferenceHistoryError);
  }
  expect(() => parse('<w:sectPrChange w:id="7" xmlns:x="urn:stella:folio:section-reference-history:1"><w:sectPr/><x:previousReferences/><x:previousReferences/></w:sectPrChange>')).toThrow(InvalidSectionReferenceHistoryError);
});
