/**
 * Property tests for the parse-warning channel.
 *
 * Folio opens input Word opens, which means normalising rather than refusing.
 * The invariant that makes that honest is not "it parsed" but "it parsed and
 * said what it changed". Over packages seeded with zero or more of each
 * normalisable defect:
 *
 *   1. The parse never throws.
 *   2. Each seeded defect yields a warning under its own code, and a clean
 *      package yields none at all.
 *   3. Each warning carries the part it happened in.
 *   4. The list is deterministic: the same bytes parse to the same ordered
 *      codes, so a host can diff two runs.
 *   5. `Document.warnings` is exactly `Document.parseWarnings` rendered, so
 *      the prose and the data cannot disagree.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { PARSE_WARNING_CODES, type ParseWarningCode } from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { formatParseWarnings } from "./parseWarningMessage";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The defect kinds this property seeds, each with the code it must report. */
const SEEDED_DEFECTS = {
  headerFooterType: PARSE_WARNING_CODES.headerFooterTypeOutsideEnum,
  onOffValue: PARSE_WARNING_CODES.unrecognisedOnOffValue,
  borderWithoutValue: PARSE_WARNING_CODES.borderWithoutValue,
  duplicateNoteId: PARSE_WARNING_CODES.duplicateNoteId,
  missingCommentId: PARSE_WARNING_CODES.missingCommentId,
  duplicateCommentId: PARSE_WARNING_CODES.duplicateCommentId,
} as const satisfies Record<string, ParseWarningCode>;

type SeededDefect = keyof typeof SEEDED_DEFECTS;

type Seed = Record<SeededDefect, number>;

const SEEDED_DEFECT_NAMES = Object.keys(SEEDED_DEFECTS) as SeededDefect[];

const seedArbitrary: fc.Arbitrary<Seed> = fc.record(
  Object.fromEntries(
    SEEDED_DEFECT_NAMES.map((name) => [name, fc.integer({ min: 0, max: 3 })]),
  ) as Record<SeededDefect, fc.Arbitrary<number>>,
);

const repeat = (count: number, build: (index: number) => string): string =>
  Array.from({ length: count }, (_unused, index) => build(index)).join("");

/**
 * The final `sectPr`, carrying the header reference and the page borders.
 *
 * A header reference needs a header part that exists, or the dangling-header
 * normaliser removes it before the type is ever compared.
 */
const sectionXml = (seed: Seed): string =>
  `<w:sectPr>${repeat(
    seed.headerFooterType,
    (index) => `<w:headerReference w:type="odd${String(index)}" r:id="rIdHdr"/>`,
  )}${
    seed.borderWithoutValue > 0
      ? `<w:pgBorders>${["top", "bottom", "left", "right"]
          .slice(0, seed.borderWithoutValue)
          .map((side) => `<w:${side} w:sz="4"/>`)
          .join("")}</w:pgBorders>`
      : ""
  }${repeat(seed.onOffValue, () => '<w:titlePg w:val="perhaps"/>')}</w:sectPr>`;

const documentXml = (seed: Seed): string =>
  `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>${sectionXml(seed)}</w:body>
</w:document>`;

const footnotesXml = (seed: Seed): string =>
  `${XML_DECLARATION}
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:t>First</w:t></w:r></w:p></w:footnote>
  ${repeat(
    seed.duplicateNoteId,
    () => '<w:footnote w:id="1"><w:p><w:r><w:t>Repeat</w:t></w:r></w:p></w:footnote>',
  )}
</w:footnotes>`;

const commentsXml = (seed: Seed): string =>
  `${XML_DECLARATION}
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="7" w:author="R"><w:p><w:r><w:t>Kept</w:t></w:r></w:p></w:comment>
  ${repeat(
    seed.duplicateCommentId,
    () => '<w:comment w:id="7" w:author="R"><w:p><w:r><w:t>Repeat</w:t></w:r></w:p></w:comment>',
  )}
  ${repeat(
    seed.missingCommentId,
    () => '<w:comment w:author="R"><w:p><w:r><w:t>No id</w:t></w:r></w:p></w:comment>',
  )}
</w:comments>`;

const docxFor = async (seed: Seed): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
  <Relationship Id="rIdHdr" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>
  <Relationship Id="rIdFn" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>
  <Relationship Id="rIdCm" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>`,
  );
  zip.file("word/document.xml", documentXml(seed));
  zip.file("word/footnotes.xml", footnotesXml(seed));
  zip.file("word/comments.xml", commentsXml(seed));
  return zip.generateAsync({ type: "arraybuffer" });
};

const codesIn = (warnings: readonly { code: ParseWarningCode }[]): ParseWarningCode[] =>
  warnings.map((warning) => warning.code);

describe("parse warnings (property)", () => {
  test(
    "every seeded normalisation reports its own code, and nothing else does",
    async () => {
      await fc.assert(
        fc.asyncProperty(seedArbitrary, async (seed) => {
          const document = await parseDocx(await docxFor(seed), { preloadFonts: false });
          const warnings = document.parseWarnings ?? [];
          const codes = codesIn(warnings);

          for (const name of SEEDED_DEFECT_NAMES) {
            const code = SEEDED_DEFECTS[name];
            expect(codes.includes(code)).toBe(seed[name] > 0);
          }
          for (const warning of warnings) {
            expect(warning.location.part.length).toBeGreaterThan(0);
            expect(warning.count).toBeGreaterThanOrEqual(1);
          }
          // Prose is rendered from the data, never written beside it.
          expect(document.warnings).toEqual(formatParseWarnings(warnings));
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(60_000),
  );

  // Non-vacuity: without this, the property above would pass just as happily
  // if nothing ever warned, because every `includes` would compare false to
  // false for a seed of zeroes.
  test("a package seeded with every defect reports every code", async () => {
    const everything: Seed = Object.fromEntries(
      SEEDED_DEFECT_NAMES.map((name) => [name, 1]),
    ) as Seed;

    const document = await parseDocx(await docxFor(everything), { preloadFonts: false });

    expect(codesIn(document.parseWarnings ?? []).toSorted()).toEqual(
      SEEDED_DEFECT_NAMES.map((name) => SEEDED_DEFECTS[name]).toSorted(),
    );
  });

  test("a package with none of the defects warns about nothing", async () => {
    const clean: Seed = Object.fromEntries(SEEDED_DEFECT_NAMES.map((name) => [name, 0])) as Seed;

    const document = await parseDocx(await docxFor(clean), { preloadFonts: false });

    expect(document.parseWarnings).toBeUndefined();
    expect(document.warnings).toBeUndefined();
  });

  test(
    "the same bytes produce the same ordered warnings",
    async () => {
      await fc.assert(
        fc.asyncProperty(seedArbitrary, async (seed) => {
          const bytes = await docxFor(seed);
          const first = await parseDocx(bytes.slice(0), { preloadFonts: false });
          const second = await parseDocx(bytes.slice(0), { preloadFonts: false });

          expect(second.parseWarnings).toEqual(first.parseWarnings);
        }),
        propertyConfig({ numRuns: 20 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
