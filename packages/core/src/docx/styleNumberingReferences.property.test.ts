/**
 * Property tests for the style tier's `w:numPr` on the `createDocx` seed path.
 *
 * `w:numId w:val="0"` is ECMA-376's "no numbering" sentinel (17.9.18, 17.9.19),
 * not a reference: on a style it cancels numbering inherited through
 * `w:basedOn`. Three invariants hold over arbitrary style packages whose every
 * style numbering is absent, the sentinel, or a defined `w:num`:
 *
 *   1. No panic — the sentinel is never read as a dangling reference.
 *   2. Sentinel preservation — a style that carried numId 0 still carries it in
 *      the written styles.xml. Dropping the `w:numPr` would hand the style its
 *      parent's numbering back, and remapping it would number the style.
 *   3. Reference preservation — a style that named a defined numId still names
 *      that same numId, and numbering.xml still defines it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { Document, Style } from "../types/document";
import { NO_NUMBERING_NUM_ID } from "./numberingReference";
import { createDocx } from "./rezip";

type StyleNumbering =
  | { kind: "none" }
  | { kind: "sentinel" }
  | { kind: "reference"; index: number };

type StylePackage = {
  /** `w:num` ids, in declaration order; every one maps to abstract num 0. */
  numIds: readonly number[];
  /** One entry per style; `basedOn` points at a lower index, so no cycles. */
  styles: readonly { basedOnIndex: number | null; numbering: StyleNumbering }[];
};

const styleIdAt = (index: number): string => `Style${String(index)}`;

/** Distinct positive `w:numId`s: 0 is reserved, so it can never collide. */
const numIdsArbitrary = fc
  .uniqueArray(fc.integer({ min: 1, max: 40 }), { minLength: 1, maxLength: 5 })
  .map((ids) => [...ids].sort((left, right) => left - right));

const stylePackageArbitrary: fc.Arbitrary<StylePackage> = numIdsArbitrary.chain((numIds) =>
  fc
    .array(
      fc.record({
        basedOnDepth: fc.integer({ min: 1, max: 4 }),
        numbering: fc.oneof(
          fc.constant<StyleNumbering>({ kind: "none" }),
          fc.constant<StyleNumbering>({ kind: "sentinel" }),
          fc
            .integer({ min: 0, max: numIds.length - 1 })
            .map<StyleNumbering>((index) => ({ kind: "reference", index })),
        ),
      }),
      { minLength: 1, maxLength: 8 },
    )
    .map((entries) => ({
      numIds,
      styles: entries.map(({ basedOnDepth, numbering }, index) => ({
        basedOnIndex: index - basedOnDepth >= 0 ? index - basedOnDepth : null,
        numbering,
      })),
    })),
);

const pPrFor = (numbering: StyleNumbering, numIds: readonly number[]): Style["pPr"] => {
  switch (numbering.kind) {
    case "none": {
      return undefined;
    }
    case "sentinel": {
      return { numPr: { numId: NO_NUMBERING_NUM_ID } };
    }
    case "reference": {
      const numId = numIds.at(numbering.index);
      if (numId === undefined) {
        throw new Error("generated numbering index is outside the generated numIds");
      }
      return { numPr: { numId, ilvl: 0 } };
    }
    default: {
      return numbering satisfies never;
    }
  }
};

const documentFor = ({ numIds, styles }: StylePackage): Document => ({
  package: {
    document: {
      finalSectionProperties: {},
      content: [{ type: "paragraph", content: [] }],
    },
    styles: {
      styles: styles.map(({ basedOnIndex, numbering }, index) => {
        const pPr = pPrFor(numbering, numIds);
        return {
          styleId: styleIdAt(index),
          type: "paragraph" as const,
          ...(basedOnIndex === null ? {} : { basedOn: styleIdAt(basedOnIndex) }),
          ...(pPr === undefined ? {} : { pPr }),
        };
      }),
    },
    numbering: {
      abstractNums: [
        { abstractNumId: 0, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
      ],
      nums: numIds.map((numId) => ({ numId, abstractNumId: 0 })),
    },
  },
});

/** The serialized `<w:style>` element for one style id. */
const styleElement = (stylesXml: string, styleId: string): string => {
  const start = stylesXml.indexOf(`<w:style w:type="paragraph" w:styleId="${styleId}">`);
  expect(start).toBeGreaterThanOrEqual(0);
  const end = stylesXml.indexOf("</w:style>", start);
  expect(end).toBeGreaterThan(start);
  return stylesXml.slice(start, end);
};

describe("createDocx style numbering (property)", () => {
  test(
    "carries every style numbering through unchanged, sentinel included",
    async () => {
      await fc.assert(
        fc.asyncProperty(stylePackageArbitrary, async (stylePackage) => {
          const zip = await JSZip.loadAsync(await createDocx(documentFor(stylePackage)));
          const stylesXml = await zip.file("word/styles.xml")?.async("text");
          const numberingXml = await zip.file("word/numbering.xml")?.async("text");
          expect(stylesXml).toBeDefined();
          expect(numberingXml).toBeDefined();

          for (const [index, { numbering }] of stylePackage.styles.entries()) {
            const element = styleElement(stylesXml ?? "", styleIdAt(index));
            switch (numbering.kind) {
              case "none": {
                expect(element).not.toContain("<w:numPr>");
                break;
              }
              case "sentinel": {
                expect(element).toContain('<w:numId w:val="0"/>');
                break;
              }
              case "reference": {
                const numId = String(stylePackage.numIds[numbering.index]);
                expect(element).toContain(`<w:numId w:val="${numId}"/>`);
                expect(numberingXml).toContain(`<w:num w:numId="${numId}">`);
                break;
              }
              default: {
                numbering satisfies never;
              }
            }
          }
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
