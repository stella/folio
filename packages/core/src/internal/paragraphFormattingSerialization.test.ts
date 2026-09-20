import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { paragraphNumberingFromSlots } from "@stll/docx-core/model";
import { serializeParagraphFormatting } from "../docx/serializer/paragraphSerializer";
import type { ParagraphFormatting } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import { modelParagraphFormattingEmission } from "./paragraphFormattingSerialization";

const COMPLETE_FORMATTING = {
  alignment: "center",
  bidi: false,
  kinsoku: true,
  overflowPunctuation: false,
  spaceBefore: 0,
  spaceAfter: 120,
  lineSpacing: 240,
  lineSpacingRule: "auto",
  snapToGrid: false,
  beforeAutospacing: true,
  afterAutospacing: false,
  spacingExplicit: { before: true, after: false },
  indentLeft: 720,
  indentRight: 0,
  indentFirstLine: -360,
  hangingIndent: true,
  borders: { bottom: { style: "single", size: 8 } },
  shading: { fill: { rgb: "E0E0E0" }, pattern: "clear" },
  tabs: [{ position: 720, alignment: "left", leader: "dot" }],
  keepNext: true,
  keepLines: false,
  widowControl: true,
  pageBreakBefore: false,
  contextualSpacing: true,
  numPr: { kind: "reference", numId: 7, ilvl: 1 },
  numPrFromStyle: { kind: "reference", numId: 7, ilvl: 1 },
  outlineLevel: { kind: "heading", level: 2 },
  styleId: "BodyText",
  frame: { dropCap: "drop", lines: 2 },
  suppressLineNumbers: false,
  suppressAutoHyphens: true,
  runProperties: { bold: true },
  runInWithNext: false,
} as const satisfies Required<ParagraphFormatting>;

const NO_OP_FORMATTINGS = [
  { runInWithNext: false },
  { hangingIndent: true },
  { borders: {} },
  { tabs: [] },
  { frame: {} },
  { runProperties: {} },
  { runProperties: { fontFamily: {} } },
  { runProperties: { styleId: "" } },
  { runProperties: { language: {} } },
  { runProperties: { color: {} } },
  { spacingExplicit: { before: true } },
  { styleId: "" },
] as const satisfies readonly ParagraphFormatting[];

describe("paragraph formatting emission model", () => {
  test("models every independent field and resolves dependent numbering provenance", () => {
    const modeled = modelParagraphFormattingEmission(COMPLETE_FORMATTING);

    expect(Object.keys(COMPLETE_FORMATTING).length).toBe(33);
    expect(modeled).toEqual({
      propertiesXml:
        '<w:pStyle w:val="BodyText"/><w:keepNext/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/><w:framePr w:dropCap="drop" w:lines="2"/><w:widowControl/><w:suppressLineNumbers w:val="0"/><w:pBdr><w:bottom w:val="single" w:sz="8"/></w:pBdr><w:shd w:val="clear" w:fill="E0E0E0"/><w:tabs><w:tab w:val="left" w:pos="720" w:leader="dot"/></w:tabs><w:suppressAutoHyphens/><w:kinsoku/><w:overflowPunct w:val="0"/><w:bidi w:val="0"/><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto" w:beforeAutospacing="1" w:afterAutospacing="0"/><w:ind w:left="720" w:right="0" w:hanging="360"/><w:contextualSpacing/><w:jc w:val="center"/><w:outlineLvl w:val="2"/>',
      paragraphMarkPropertiesInnerXml: "<w:b/>",
    });
  });

  test.each(NO_OP_FORMATTINGS)("normalizes emission-empty formatting %#", (formatting) => {
    expect(Object.keys(formatting).length).toBeGreaterThan(0);
    expect(modelParagraphFormattingEmission(formatting)).toEqual({});
    expect(serializeParagraphFormatting(formatting)).toBe("");
  });

  test("normalizes an explicit non-hanging indent to the default first-line instruction", () => {
    const firstLine = { indentFirstLine: 120 } satisfies ParagraphFormatting;
    const explicitNonHanging = {
      indentFirstLine: 120,
      hangingIndent: false,
    } satisfies ParagraphFormatting;

    expect(modelParagraphFormattingEmission(explicitNonHanging)).toEqual(
      modelParagraphFormattingEmission(firstLine),
    );
    expect(serializeParagraphFormatting(explicitNonHanging)).toBe(
      serializeParagraphFormatting(firstLine),
    );
  });

  test.each([{ fontFamily: {} }, { styleId: "" }, { language: {} }, { color: {} }] as const)(
    "normalizes nested empty run formatting %# alongside emitted formatting",
    (empty) => {
      const formatting = { runProperties: { bold: true, ...empty } } satisfies ParagraphFormatting;
      const bold = { runProperties: { bold: true } } satisfies ParagraphFormatting;

      expect(modelParagraphFormattingEmission(formatting)).toEqual(
        modelParagraphFormattingEmission(bold),
      );
      expect(serializeParagraphFormatting(formatting)).toBe(serializeParagraphFormatting(bold));
    },
  );

  test("modeled equality agrees with fallback XML across generated equivalent states", () => {
    const formattingArbitrary = fc.record({
      alignment: fc.constantFrom("left", "center", "both"),
      keepNext: fc.boolean(),
      numPr: fc.record({
        kind: fc.constant("reference" as const),
        numId: fc.integer({ min: 1, max: 100 }),
        ilvl: fc.integer({ min: 0, max: 8 }),
      }),
      spaceAfter: fc.integer({ min: 0, max: 2_000 }),
    });

    fc.assert(
      fc.property(
        formattingArbitrary,
        fc.subarray(NO_OP_FORMATTINGS, { minLength: 1 }),
        (formatting, noOps) => {
          const equivalentFormatting = Object.assign({ ...formatting }, ...noOps);
          expect(canonicalJson(equivalentFormatting)).not.toBe(canonicalJson(formatting));
          expect(modelParagraphFormattingEmission(equivalentFormatting)).toEqual(
            modelParagraphFormattingEmission(formatting),
          );
          expect(serializeParagraphFormatting(equivalentFormatting)).toBe(
            serializeParagraphFormatting(formatting),
          );

          const changedFormatting = {
            ...equivalentFormatting,
            keepNext: !formatting.keepNext,
          };
          expect(modelParagraphFormattingEmission(changedFormatting)).not.toEqual(
            modelParagraphFormattingEmission(formatting),
          );
          expect(serializeParagraphFormatting(changedFormatting)).not.toBe(
            serializeParagraphFormatting(formatting),
          );
        },
      ),
      { numRuns: 128 },
    );
  });

  test.each([
    ["absent", undefined, undefined, {}],
    [
      "cancelled",
      { kind: "none" },
      undefined,
      { propertiesXml: '<w:numPr><w:numId w:val="0"/></w:numPr>' },
    ],
    [
      "direct",
      { kind: "reference", numId: 7, ilvl: 1 },
      undefined,
      { propertiesXml: '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>' },
    ],
    [
      "style-sourced",
      { kind: "reference", numId: 7, ilvl: 1 },
      { kind: "reference", numId: 7, ilvl: 1 },
      {},
    ],
    [
      "implicit style level zero",
      { kind: "reference", numId: 7 },
      { kind: "reference", numId: 7, ilvl: 0 },
      {},
    ],
    [
      "level stated without an id",
      { kind: "levelOnly", ilvl: 2 },
      undefined,
      { propertiesXml: '<w:numPr><w:ilvl w:val="2"/></w:numPr>' },
    ],
    [
      "changed from style",
      { kind: "reference", numId: 7, ilvl: 1 },
      { kind: "reference", numId: 8, ilvl: 1 },
      { propertiesXml: '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>' },
    ],
  ] as const)(
    "resolves %s numbering into emitted numPr",
    (_name, numPr, numPrFromStyle, expected) => {
      expect(modelParagraphFormattingEmission({ numPr, numPrFromStyle })).toEqual(expected);
    },
  );

  test("generated models reconstruct the exact fallback XML", () => {
    fc.assert(
      fc.property(
        fc.option(fc.integer({ min: 0, max: 100 }), { nil: undefined }),
        fc.option(fc.integer({ min: 0, max: 8 }), { nil: undefined }),
        fc.boolean(),
        (numId, ilvl, styleSourced) => {
          const numPr = paragraphNumberingFromSlots({ numId, ilvl });
          const formatting = {
            alignment: "both",
            numPr,
            ...(styleSourced ? { numPrFromStyle: numPr } : {}),
            spacingExplicit: { before: true },
          } as const satisfies ParagraphFormatting;

          const modeled = modelParagraphFormattingEmission(formatting);
          const innerXml = `${modeled.propertiesXml ?? ""}${
            modeled.paragraphMarkPropertiesInnerXml
              ? `<w:rPr>${modeled.paragraphMarkPropertiesInnerXml}</w:rPr>`
              : ""
          }`;
          expect(serializeParagraphFormatting(formatting)).toBe(
            innerXml ? `<w:pPr>${innerXml}</w:pPr>` : "",
          );
        },
      ),
      { numRuns: 128 },
    );
  });
});
