/**
 * Property tests for the style tier's `w:numPr` on the `createDocx` seed path.
 *
 * `w:numId w:val="0"` is ECMA-376's "no numbering" sentinel (17.9.18, 17.9.19),
 * not a reference: on a style it cancels numbering inherited through
 * `w:basedOn`. Three invariants hold over arbitrary style packages Folio built,
 * whose every style numbering is absent, the sentinel, or a defined `w:num`:
 *
 *   1. No panic — the sentinel is never read as a dangling reference.
 *   2. Sentinel preservation — a style that carried numId 0 still carries it in
 *      the written styles.xml. Dropping the `w:numPr` would hand the style its
 *      parent's numbering back, and remapping it would number the style.
 *   3. Reference preservation — a style that named a defined numId still names
 *      that same numId, and numbering.xml still defines it.
 *
 * A source file can also name a `w:num` the numbering part never defines, which
 * Word opens and renders unnumbered. The second property adds that fourth case
 * and runs the route a user's styles actually travel, parse to style set to
 * `createDocx`: the dangling style must come out unnumbered rather than
 * inheriting its parent's list, the save must not panic, and the parse must say
 * so once per style.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { extractDocumentStyleSet } from "../style-sets/extract";
import type { Document, Style } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { NO_NUMBERING_NUM_ID, NO_PARAGRAPH_NUMBERING } from "./numberingReference";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createDocx } from "./rezip";

type StyleNumbering =
  | { kind: "none" }
  | { kind: "sentinel" }
  | { kind: "reference"; numId: number }
  | { kind: "dangling" };

type StylePackage = {
  /** `w:num` ids, in declaration order; every one maps to abstract num 0. */
  numIds: readonly number[];
  /** One entry per style; `basedOn` points at a lower index, so no cycles. */
  styles: readonly { basedOnIndex: number | null; numbering: StyleNumbering }[];
};

/** What Folio itself can put in the model, once the parse boundary is passed. */
const FOLIO_BUILT_KINDS = ["none", "sentinel", "reference"] as const;
const SOURCE_KINDS = [...FOLIO_BUILT_KINDS, "dangling"] as const;

const styleIdAt = (index: number): string => `Style${String(index)}`;

/** Outside the generated `w:num` ids, so it can never accidentally resolve. */
const danglingNumIdAt = (index: number): number => 100 + index;

/** Distinct positive `w:numId`s: 0 is reserved, so it can never collide. */
const numIdsArbitrary = fc
  .uniqueArray(fc.integer({ min: 1, max: 40 }), { minLength: 1, maxLength: 5 })
  .map((ids) => [...ids].sort((left, right) => left - right));

const numberingArbitrary = (
  numIds: readonly number[],
  kinds: readonly StyleNumbering["kind"][],
): fc.Arbitrary<StyleNumbering> =>
  fc.oneof(
    ...kinds.map((kind) =>
      kind === "reference"
        ? fc.constantFrom(...numIds).map<StyleNumbering>((numId) => ({ kind: "reference", numId }))
        : fc.constant<StyleNumbering>({ kind }),
    ),
  );

const stylePackageArbitrary = (
  kinds: readonly StyleNumbering["kind"][],
): fc.Arbitrary<StylePackage> =>
  numIdsArbitrary.chain((numIds) =>
    fc
      .array(
        fc.record({
          basedOnDepth: fc.integer({ min: 1, max: 4 }),
          numbering: numberingArbitrary(numIds, kinds),
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

/**
 * The numbering a style ends up with. `w:numId` and `w:ilvl` inherit per field
 * through `w:basedOn`, so a style that states no `w:numPr` carries its parent's.
 */
const effectiveNumbering = (styles: StylePackage["styles"], index: number): StyleNumbering => {
  const entry = styles[index];
  if (entry === undefined) {
    throw new Error("style index is outside the generated package");
  }
  if (entry.numbering.kind !== "none" || entry.basedOnIndex === null) {
    return entry.numbering;
  }
  return effectiveNumbering(styles, entry.basedOnIndex);
};

const numIdFor = (numbering: StyleNumbering, index: number): number => {
  switch (numbering.kind) {
    case "sentinel": {
      return NO_NUMBERING_NUM_ID;
    }
    case "dangling": {
      return danglingNumIdAt(index);
    }
    case "reference": {
      return numbering.numId;
    }
    case "none": {
      throw new Error("a style that states no numbering has no w:numId");
    }
    default: {
      return numbering satisfies never;
    }
  }
};

const pPrFor = (numbering: StyleNumbering, index: number): Style["pPr"] => {
  if (numbering.kind === "none") {
    return undefined;
  }
  const numId = numIdFor(numbering, index);
  return numbering.kind === "sentinel"
    ? { numPr: NO_PARAGRAPH_NUMBERING }
    : { numPr: { kind: "reference" as const, numId, ilvl: 0 } };
};

const documentFor = ({ numIds, styles }: StylePackage): Document => ({
  package: {
    document: {
      finalSectionProperties: {},
      content: [{ type: "paragraph", content: [] }],
    },
    styles: {
      styles: styles.map(({ basedOnIndex, numbering }, index) => {
        const pPr = pPrFor(numbering, index);
        return {
          styleId: styleIdAt(index),
          type: "paragraph",
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

const UNNUMBERED_STYLE_PREFIX = "Unnumbered style ";

const unnumberedStyleWarning = (styleId: string): string =>
  `${UNNUMBERED_STYLE_PREFIX}"${styleId}" whose numbering definition is missing.`;

const styleIdsWhere = (
  { styles }: StylePackage,
  predicate: (style: StylePackage["styles"][number], index: number) => boolean,
): string[] =>
  styles.flatMap((style, index) => (predicate(style, index) ? [styleIdAt(index)] : []));

/** The serialized `<w:style>` element for one style id. */
const styleElement = (stylesXml: string, styleId: string): string => {
  const start = stylesXml.indexOf(`w:styleId="${styleId}">`);
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
        fc.asyncProperty(stylePackageArbitrary(FOLIO_BUILT_KINDS), async (stylePackage) => {
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
                const numId = String(numbering.numId);
                expect(element).toContain(`<w:numId w:val="${numId}"/>`);
                expect(numberingXml).toContain(`<w:num w:numId="${numId}">`);
                break;
              }
              case "dangling": {
                throw new Error("a Folio-built package cannot carry a dangling reference");
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

describe("source style numbering through parse and extraction (property)", () => {
  test(
    "unnumbers every dangling style once and saves without a panic",
    async () => {
      await fc.assert(
        fc.asyncProperty(stylePackageArbitrary(SOURCE_KINDS), async (stylePackage) => {
          const document = await parseDocx(await docxFor(stylePackage), { preloadFonts: false });

          // Sound and complete, once each: a style that states a dangling
          // reference is reported, no style is reported twice, and nothing is
          // reported whose numbering resolves. A style that only inherits a
          // dangling reference shares its parent's properties, so whether it is
          // reported on its own is an implementation detail, not an invariant.
          const reported = (document.warnings ?? []).filter((warning) =>
            warning.startsWith(UNNUMBERED_STYLE_PREFIX),
          );
          expect(reported).toEqual([...new Set(reported)]);
          expect(reported).toEqual(
            expect.arrayContaining(
              styleIdsWhere(stylePackage, ({ numbering }) => numbering.kind === "dangling").map(
                unnumberedStyleWarning,
              ),
            ),
          );
          const reportable = styleIdsWhere(
            stylePackage,
            (_style, index) => effectiveNumbering(stylePackage.styles, index).kind === "dangling",
          ).map(unnumberedStyleWarning);
          for (const warning of reported) {
            expect(reportable).toContain(warning);
          }

          const styleSet = extractDocumentStyleSet(document, { name: "Extracted" });
          const zip = await JSZip.loadAsync(await createDocx(createEmptyDocument({ styleSet })));
          const stylesXml = (await zip.file("word/styles.xml")?.async("text")) ?? "";
          const numberingXml = (await zip.file("word/numbering.xml")?.async("text")) ?? "";

          for (const [index] of stylePackage.styles.entries()) {
            const numbering = effectiveNumbering(stylePackage.styles, index);
            const element = styleElement(stylesXml, styleIdAt(index));
            if (numbering.kind === "none") {
              expect(element).not.toContain("<w:numPr>");
              continue;
            }
            // A dangling reference comes out as the sentinel: unnumbered, and
            // not reopened to the parent's list by a deleted <w:numPr>.
            const numId = numbering.kind === "dangling" ? 0 : numIdFor(numbering, index);
            expect(element).toContain(`<w:numId w:val="${String(numId)}"/>`);
            if (numId !== NO_NUMBERING_NUM_ID) {
              expect(numberingXml).toContain(`<w:num w:numId="${String(numId)}">`);
            }
          }
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const stylesXmlFor = ({ styles }: StylePackage): string =>
  `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  ${styles
    .map(({ basedOnIndex, numbering }, index) => {
      const basedOn =
        basedOnIndex === null ? "" : `<w:basedOn w:val="${styleIdAt(basedOnIndex)}"/>`;
      const pPr =
        numbering.kind === "none"
          ? ""
          : `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${String(numIdFor(numbering, index))}"/></w:numPr></w:pPr>`;
      return `<w:style w:type="paragraph" w:styleId="${styleIdAt(index)}">${basedOn}${pPr}</w:style>`;
    })
    .join("")}
</w:styles>`;

const numberingXmlFor = ({ numIds }: StylePackage): string =>
  `${XML_DECLARATION}
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
  </w:abstractNum>
  ${numIds
    .map((numId) => `<w:num w:numId="${String(numId)}"><w:abstractNumId w:val="0"/></w:num>`)
    .join("")}
</w:numbering>`;

/** A minimal source package carrying the generated styles and numbering. */
const docxFor = async (stylePackage: StylePackage): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.numbering}" Target="numbering.xml"/>
</Relationships>`,
  );
  zip.file("word/styles.xml", stylesXmlFor(stylePackage));
  zip.file("word/numbering.xml", numberingXmlFor(stylePackage));
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body>
</w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};
