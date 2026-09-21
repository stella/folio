/**
 * A compiled `w:rPr` writes its children in the schema's order, whatever the
 * source states.
 *
 * This package holds a second `w:rPr` writer — the one the legal-source
 * compiler and the build-from-scratch export share — and it had grown an order
 * of its own, emitting `w:highlight`, `w:sz` and `w:szCs` ahead of `w:rFonts`.
 * `{fontFamily, fontSize}` is the smallest source that shows it: folio-core's
 * serializer writes `<w:rFonts/><w:sz/>` for that run and this one wrote
 * `<w:sz/><w:rFonts/>`.
 *
 * The order is not a validity rule here. `EG_RPrBase` is an `xsd:choice`
 * referenced `maxOccurs="unbounded"`, so both spellings are schema-valid and
 * the repository's own validator reports nothing for either; what the order
 * buys is one canonical form, the one Word writes, from both of folio's
 * writers. So the property is stated against the generated list rather than
 * against the validator, which cannot see this class of defect at all.
 *
 * The sample map is `satisfies Record<keyof TextFormatting, …>`: a field the
 * model gains cannot reach this property without somebody stating what it
 * looks like, and the subset generator then puts it in front of every other
 * field the writer emits.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import type { Document } from "../model/document";
import type { TextFormatting } from "../model/formatting";
import { SEQUENCE_CHILDREN } from "../schema/sequenceChildren";
import { serializeDocumentToDocx } from "./docx";

/** One minimal `TextFormatting` per model field, each stating that field. */
const SAMPLES = {
  bold: { bold: true },
  boldCs: { boldCs: false },
  italic: { italic: true },
  italicCs: { italicCs: false },
  underline: { underline: { style: "single" } },
  strike: { strike: true },
  doubleStrike: { doubleStrike: true },
  vertAlign: { vertAlign: "superscript" },
  smallCaps: { smallCaps: true },
  allCaps: { allCaps: true },
  hidden: { hidden: true },
  noProof: { noProof: true },
  color: { color: { rgb: "1F4E79" } },
  highlight: { highlight: "yellow" },
  shading: { shading: { pattern: "clear", fill: { rgb: "D9D9D9" } } },
  fontSize: { fontSize: 24 },
  fontSizeCs: { fontSizeCs: 26 },
  fontFamily: { fontFamily: { ascii: "Georgia", hAnsi: "Georgia", cs: "Amiri" } },
  language: { language: { val: "cs-CZ" } },
  spacing: { spacing: 20 },
  position: { position: 6 },
  scale: { scale: 110 },
  kerning: { kerning: 18 },
  effect: { effect: "shimmer" },
  emphasisMark: { emphasisMark: "dot" },
  emboss: { emboss: true },
  imprint: { imprint: true },
  outline: { outline: true },
  shadow: { shadow: true },
  rtl: { rtl: true },
  cs: { cs: true },
  styleId: { styleId: "Emphasis" },
  // Keeps this map total; the compiled writer has no capture channel, so this
  // sample contributes no child to its emitted `w:rPr`.
  preserved: { preserved: { children: [{ index: 0, xml: "<w:bdr/>" }] } },
} as const satisfies Record<keyof TextFormatting, TextFormatting>;

const FIELDS = Object.keys(SAMPLES) as (keyof typeof SAMPLES)[];

const documentWith = (formatting: TextFormatting): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "x" }], formatting }],
        },
      ],
    },
  },
});

const readDocumentXml = async (buf: ArrayBuffer): Promise<string> => {
  const xml = await (await JSZip.loadAsync(buf)).file("word/document.xml")?.async("string");
  if (xml === undefined) {
    throw new Error("word/document.xml missing from serialized docx");
  }
  return xml;
};

/**
 * Where each child of the first `w:rPr` sits in the generated order.
 *
 * The delimiter matters: `w:b` is a prefix of `w:bCs`, so matching on the name
 * alone reports one child at another's place.
 */
const declaredPositions = (documentXml: string): number[] => {
  const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/u.exec(documentXml)?.[1] ?? "";
  const declared: readonly string[] = SEQUENCE_CHILDREN["run-properties"];
  const positions: number[] = [];
  for (const [, name] of rPr.matchAll(/<w:([A-Za-z0-9]+)[\s/>]/gu)) {
    const at = declared.indexOf(name ?? "");
    expect(at).toBeGreaterThan(-1);
    positions.push(at);
  }
  return positions;
};

test("a compiled run property set is written in the schema's child order", async () => {
  await fc.assert(
    fc.asyncProperty(fc.subarray(FIELDS, { minLength: 2 }), async (fields) => {
      const formatting = Object.assign({}, ...fields.map((field) => SAMPLES[field]));
      const positions = declaredPositions(
        await readDocumentXml(await serializeDocumentToDocx(documentWith(formatting))),
      );

      expect(positions).toEqual([...positions].toSorted((left, right) => left - right));
    }),
    propertyConfig({ numRuns: 200 }),
  );
});

test("a font and a size come out in the order the schema declares them", async () => {
  // The smallest source that failed before: `w:sz` was written ahead of
  // `w:rFonts`, so the same run compiled by this package and written by
  // folio-core's serializer came out in two different orders.
  const documentXml = await readDocumentXml(
    await serializeDocumentToDocx(
      documentWith({ fontFamily: { ascii: "Georgia" }, fontSize: 24, highlight: "yellow" }),
    ),
  );

  expect(documentXml).toContain(
    '<w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="24"/><w:highlight w:val="yellow"/></w:rPr>',
  );
});
