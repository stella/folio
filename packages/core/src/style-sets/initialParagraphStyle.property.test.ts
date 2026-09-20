/**
 * Property tests for the initial paragraph style of an extracted style set.
 *
 * `w:styleId` is opaque: nothing obliges a document to call its default
 * paragraph style `Normal`, and localized or generated packages routinely do
 * not (`Standard`, `Normln`, `style0`). Resolution follows ECMA-376 17.7.4.17
 * instead: the paragraph style flagged `w:default="1"` is the default, the last
 * one wins where several are flagged, and a package that flags none leaves the
 * consumer's built-in Normal to apply.
 *
 * Over arbitrary style packages, with zero or one flagged default under
 * arbitrary ids:
 *
 *   1. No panic, and the set names a paragraph style it actually contains.
 *   2. A flagged default is the one chosen.
 *   3. Nothing the source declared is dropped from the set.
 *   4. The set builds a package, which is the route a corpus file travels.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { createDocx } from "../docx/rezip";
import type { Document, Style } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { extractDocumentStyleSet } from "./extract";

setDefaultTimeout(propertyTestTimeout(30_000));

type GeneratedStyle = {
  styleId: string;
  type: "paragraph" | "character";
  /** `w:name`, which is where Word records a built-in style's identity. */
  name: string | undefined;
};

type StylePackage = {
  styles: readonly GeneratedStyle[];
  /** Index of the style flagged `w:default="1"`, if the package flags one. */
  defaultIndex: number | null;
};

/**
 * Ids real producers use, plus an arbitrary one.
 *
 * `Normal` is in the pool on purpose: the English-Word case must keep working,
 * and it is the id the minted fallback wants, so a generated package that
 * already holds it exercises the collision path.
 */
const styleIdArbitrary = fc.oneof(
  fc.constantFrom("Normal", "Standard", "Normln", "style0", "Default", "BodyText", "berschrift1"),
  fc.stringMatching(/^[A-Za-z][A-Za-z0-9]{0,15}$/u),
);

/** `w:name` is absent, the built-in name, or a name of the producer's own. */
const styleNameArbitrary = fc.oneof(
  fc.constant(undefined),
  fc.constantFrom("Normal", "Default", "Text body", "Heading"),
  fc.stringMatching(/^[A-Za-z][A-Za-z0-9 ]{0,15}$/u),
);

const stylePackageArbitrary: fc.Arbitrary<StylePackage> = fc
  .uniqueArray(
    fc.record({
      styleId: styleIdArbitrary,
      type: fc.constantFrom<GeneratedStyle["type"]>("paragraph", "character"),
      name: styleNameArbitrary,
    }),
    { minLength: 0, maxLength: 8, selector: (style) => style.styleId },
  )
  .chain((styles) => {
    const paragraphIndexes = styles.flatMap((style, index) =>
      style.type === "paragraph" ? [index] : [],
    );
    return fc.record({
      styles: fc.constant(styles),
      defaultIndex:
        paragraphIndexes.length === 0
          ? fc.constant(null)
          : fc.oneof(fc.constant(null), fc.constantFrom(...paragraphIndexes)),
    });
  });

const documentFor = ({ styles, defaultIndex }: StylePackage): Document => ({
  package: {
    document: {
      finalSectionProperties: {},
      content: [{ type: "paragraph", content: [] }],
    },
    styles: {
      styles: styles.map(({ styleId, type, name }, index) => ({
        styleId,
        type,
        ...(name === undefined ? {} : { name }),
        ...(index === defaultIndex ? { default: true } : {}),
      })),
    },
  },
});

const styleWithId = (styles: readonly Style[], styleId: string): Style | undefined =>
  styles.find((style) => style.styleId === styleId);

describe("extracted initial paragraph style (property)", () => {
  test("is always a paragraph style the set contains", () => {
    fc.assert(
      fc.property(stylePackageArbitrary, (stylePackage) => {
        const styleSet = extractDocumentStyleSet(documentFor(stylePackage), { name: "Extracted" });

        const initial = styleWithId(styleSet.styles.styles, styleSet.initialParagraphStyleId);
        expect(initial?.type).toBe("paragraph");

        const { styles, defaultIndex } = stylePackage;
        if (defaultIndex !== null) {
          expect(styleSet.initialParagraphStyleId).toBe(styles[defaultIndex]?.styleId);
        }
        // Minting a default may add a style; it may never remove one.
        for (const style of styles) {
          expect(styleWithId(styleSet.styles.styles, style.styleId)?.type).toBe(style.type);
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test(
    "builds a package, whatever the source declared",
    async () => {
      await fc.assert(
        fc.asyncProperty(stylePackageArbitrary, async (stylePackage) => {
          const styleSet = extractDocumentStyleSet(documentFor(stylePackage), {
            name: "Extracted",
          });
          await createDocx(createEmptyDocument({ styleSet }));
        }),
        propertyConfig({ numRuns: 30 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
