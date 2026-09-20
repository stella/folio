/**
 * A programmatic content control writes the same `<w:sdtPr>` wherever it sits.
 *
 * Block and inline SDTs used to synthesize their fallback properties from two
 * near-identical builders, and the two had drifted on the only two elements
 * where a spelling decision exists: the block one wrote `w:id` for a
 * non-integer and wrote `<w:lock w:val="unlocked"/>`, the inline one wrote
 * neither. The same control then round-tripped differently depending on where
 * it was placed. Compare the serializers' own output, over every `sdtType` and
 * over the properties that used to differ.
 */

import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  BlockSdt,
  InlineSdt,
  Paragraph,
  SdtProperties,
} from "../../types/document";
import { serializeBlockSdt } from "./blockSdtSerializer";
import { serializeParagraph } from "./paragraphSerializer";

const SDT_TYPES = [
  "richText",
  "plainText",
  "date",
  "dropdown",
  "comboBox",
  "checkbox",
  "picture",
  "buildingBlockGallery",
  "group",
  "unknown",
] as const satisfies readonly SdtProperties["sdtType"][];

const LOCKS = [
  undefined,
  "unlocked",
  "sdtLocked",
  "contentLocked",
  "sdtContentLocked",
] as const satisfies readonly (SdtProperties["lock"] | undefined)[];

const sdtPrOf = (xml: string): string => {
  const start = xml.indexOf("<w:sdtPr>");
  const end = xml.indexOf("</w:sdtPr>");
  if (start === -1 || end === -1) return xml;
  return xml.slice(start, end + "</w:sdtPr>".length);
};

const asBlock = (properties: SdtProperties): BlockSdt => ({
  type: "blockSdt",
  properties,
  content: [{ type: "paragraph", content: [] }],
});

const asParagraph = (properties: SdtProperties): Paragraph => {
  const inline: InlineSdt = { type: "inlineSdt", properties, content: [] };
  return { type: "paragraph", content: [inline] };
};

const blockSdtPr = (properties: SdtProperties): string =>
  sdtPrOf(serializeBlockSdt(asBlock(properties), (_block: BlockContent) => "<w:p/>"));

const inlineSdtPr = (properties: SdtProperties): string =>
  sdtPrOf(serializeParagraph(asParagraph(properties)));

describe("fallback w:sdtPr parity", () => {
  for (const sdtType of SDT_TYPES) {
    test(`block and inline agree on a ${sdtType} control`, () => {
      const properties = {
        sdtType,
        id: 42,
        alias: "Party name",
        tag: "party-name",
        placeholder: "PartyPlaceholder",
        showingPlaceholder: true,
        dateFormat: "d MMMM yyyy",
        dateValueISO: "2026-06-02T00:00:00Z",
        listItems: [{ displayText: "Buyer", value: "buyer" }],
        checked: true,
      } as const satisfies SdtProperties;
      expect(inlineSdtPr(properties)).toBe(blockSdtPr(properties));
    });
  }

  for (const lock of LOCKS) {
    test(`block and inline agree on lock=${String(lock)}`, () => {
      const properties: SdtProperties = {
        sdtType: "richText",
        ...(lock === undefined ? {} : { lock }),
      };
      const serialized = blockSdtPr(properties);
      expect(inlineSdtPr(properties)).toBe(serialized);
      // `w:lock` is absent by default, so `unlocked` is the absence.
      expect(serialized.includes("<w:lock")).toBe(lock !== undefined && lock !== "unlocked");
    });
  }

  test("w:id is written only for a decimal number", () => {
    // `ST_DecimalNumber` is xsd:integer; a fractional id has no spelling.
    const integer: SdtProperties = { sdtType: "richText", id: 7 };
    expect(blockSdtPr(integer)).toContain('<w:id w:val="7"/>');
    expect(inlineSdtPr(integer)).toBe(blockSdtPr(integer));

    const fractional = { sdtType: "richText", id: 7.5 } as const satisfies SdtProperties;
    expect(blockSdtPr(fractional)).not.toContain("<w:id");
    expect(inlineSdtPr(fractional)).toBe(blockSdtPr(fractional));
  });
});
