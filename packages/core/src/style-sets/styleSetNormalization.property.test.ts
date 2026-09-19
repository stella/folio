/**
 * Property tests for the boundary an external `DocumentStyleSet` crosses.
 *
 * A style set is persisted as JSON and handed back later, so one written
 * before folio learned to repair a defect still carries it: a style naming a
 * `w:num` the set never defines, two styles under one id, an initial paragraph
 * style the set does not contain. Over arbitrary v1 sets carrying any mix of
 * those:
 *
 *   1. `createEmptyDocument` never panics, and neither does `createDocx`.
 *   2. The written package parses back, and the styles that survived are the
 *      styles it carries.
 *   3. The document's initial paragraph reads a style the package defines.
 *   4. Each repair is reported, so a host learns the set it stored was fixed.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { NO_NUMBERING_NUM_ID } from "../docx/numberingReference";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { DOCUMENT_STYLE_SET_VERSION, type DocumentStyleSet } from "./types";

/** A numbering slot: absent, the "no numbering" sentinel, defined, or dangling. */
const NUMBERING_KINDS = ["none", "sentinel", "defined", "dangling"] as const;
type NumberingKind = (typeof NUMBERING_KINDS)[number];

const DEFINED_NUM_ID = 1;
const DANGLING_NUM_ID = 77;

type GeneratedStyleSet = {
  /** One entry per style, in declaration order. */
  styles: readonly { styleId: string; numbering: NumberingKind; isDefault: boolean }[];
  /** The id the set names as its initial paragraph style. */
  initialParagraphStyleId: string;
  /** Repeat the first style's id on a later style. */
  duplicateFirstStyleId: boolean;
};

const styleIdArbitrary = fc.constantFrom("Normal", "Standard", "Body", "style0", "Nadpis1");

const generatedStyleSetArbitrary: fc.Arbitrary<GeneratedStyleSet> = fc
  .uniqueArray(
    fc.record({
      styleId: styleIdArbitrary,
      numbering: fc.constantFrom(...NUMBERING_KINDS),
      isDefault: fc.boolean(),
    }),
    { minLength: 1, maxLength: 5, selector: (style) => style.styleId },
  )
  .chain((styles) =>
    fc.record({
      styles: fc.constant(styles),
      // "Missing" is in the pool on purpose: a set naming a style it does not
      // contain is exactly the shape that used to reach a panic.
      initialParagraphStyleId: fc.oneof(
        fc.constantFrom(...styles.map((style) => style.styleId)),
        fc.constant("AStyleTheSetDoesNotHave"),
      ),
      duplicateFirstStyleId: fc.boolean(),
    }),
  );

const numPrFor = (numbering: NumberingKind) => {
  switch (numbering) {
    case "none":
      return undefined;
    case "sentinel":
      return { numPr: { numId: NO_NUMBERING_NUM_ID } };
    case "defined":
      return { numPr: { numId: DEFINED_NUM_ID, ilvl: 0 } };
    case "dangling":
      return { numPr: { numId: DANGLING_NUM_ID, ilvl: 0 } };
    default: {
      return numbering satisfies never;
    }
  }
};

const styleSetFor = (generated: GeneratedStyleSet): DocumentStyleSet => {
  const styles = generated.styles.map(({ styleId, numbering, isDefault }) => {
    const pPr = numPrFor(numbering);
    return {
      styleId,
      type: "paragraph" as const,
      ...(isDefault ? { default: true } : {}),
      ...(pPr === undefined ? {} : { pPr }),
    };
  });
  const first = styles.at(0);
  return {
    version: DOCUMENT_STYLE_SET_VERSION,
    name: "Persisted set",
    initialParagraphStyleId: generated.initialParagraphStyleId,
    styles: {
      styles:
        generated.duplicateFirstStyleId && first
          ? [...styles, { ...first, name: "A repeat of the first id" }]
          : styles,
    },
    numbering: {
      abstractNums: [
        { abstractNumId: 0, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
      ],
      nums: [{ numId: DEFINED_NUM_ID, abstractNumId: 0 }],
    },
  };
};

describe("document style set boundary (property)", () => {
  test(
    "any v1 set builds a package that parses back",
    async () => {
      await fc.assert(
        fc.asyncProperty(generatedStyleSetArbitrary, async (generated) => {
          const styleSet = styleSetFor(generated);
          const before = structuredClone(styleSet);

          const document = createEmptyDocument({ styleSet });
          const reparsed = await parseDocx(await createDocx(document), { preloadFonts: false });

          // The caller's value is theirs; normalisation works on a copy.
          expect(styleSet).toEqual(before);

          const writtenIds = (reparsed.package.styles?.styles ?? []).map((style) => style.styleId);
          expect(writtenIds).toEqual([...new Set(writtenIds)]);
          expect(new Set(writtenIds)).toEqual(
            new Set((document.package.styles?.styles ?? []).map((style) => style.styleId)),
          );

          const firstBlock = reparsed.package.document.content.at(0);
          expect(firstBlock?.type).toBe("paragraph");
          const styleId =
            firstBlock?.type === "paragraph" ? firstBlock.formatting?.styleId : undefined;
          expect(writtenIds).toContain(styleId);
        }),
        propertyConfig({ numRuns: 50 }),
      );
    },
    propertyTestTimeout(90_000),
  );

  test("a set naming a style it does not contain is repaired and reported", () => {
    const document = createEmptyDocument({
      styleSet: styleSetFor({
        styles: [{ styleId: "Standard", numbering: "none", isDefault: true }],
        initialParagraphStyleId: "AStyleTheSetDoesNotHave",
        duplicateFirstStyleId: false,
      }),
    });

    expect(document.parseWarnings?.map((warning) => warning.code)).toEqual([
      PARSE_WARNING_CODES.styleSetInitialStyleMissing,
    ]);
    const firstBlock = document.package.document.content.at(0);
    expect(firstBlock?.type === "paragraph" ? firstBlock.formatting?.styleId : undefined).toBe(
      "Standard",
    );
  });

  test("a set repeating a style id keeps the first and reports the rest", () => {
    const document = createEmptyDocument({
      styleSet: styleSetFor({
        styles: [{ styleId: "Normal", numbering: "none", isDefault: true }],
        initialParagraphStyleId: "Normal",
        duplicateFirstStyleId: true,
      }),
    });

    expect(document.package.styles?.styles.map((style) => style.styleId)).toEqual(["Normal"]);
    expect(document.parseWarnings?.map((warning) => warning.code)).toEqual([
      PARSE_WARNING_CODES.styleSetDuplicateStyleId,
    ]);
  });

  test("a set whose style numbering is dangling is unnumbered and reported", () => {
    const document = createEmptyDocument({
      styleSet: styleSetFor({
        styles: [{ styleId: "Normal", numbering: "dangling", isDefault: true }],
        initialParagraphStyleId: "Normal",
        duplicateFirstStyleId: false,
      }),
    });

    expect(document.package.styles?.styles.at(0)?.pPr?.numPr?.numId).toBe(NO_NUMBERING_NUM_ID);
    expect(document.parseWarnings?.map((warning) => warning.code)).toEqual([
      PARSE_WARNING_CODES.unnumberedStyle,
    ]);
  });
});
