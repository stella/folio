/**
 * `compare(x, x)` is the identity.
 *
 * For every document the engine can parse, comparing it with itself must
 * report no change, and accepting or rejecting the result must reproduce it.
 * The engine had two ways to break that law on documents Word opens without
 * complaint: an inline atom it could not detach from its package aborted the
 * alignment of the whole story, and a package with no `word/styles.xml` was
 * refused outright because there was no part to read a default paragraph
 * style from.
 *
 * The arbitrary builds packages from the constructs behind both: a drawing
 * that is not a picture and so carries no media, a picture that is, tables,
 * and a package that defines no styles at all.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { compareDocx } from "./compare";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const COMPARE_OPTIONS = { author: "corpus-gate", timestamp: "2026-01-01T00:00:00.000Z" } as const;

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  `xmlns:r="${OFFICE_RELATIONSHIPS}"`,
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

const IMAGE_RID = "rId2";

/** A one-pixel PNG, so a picture has real bytes behind it. */
const PNG_BYTES = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  ),
  (character) => character.codePointAt(0) ?? 0,
);

const PICTURE = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${IMAGE_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

/** A `wp:inline` with no `a:graphic`: what a chart or an OLE frame reduces to. */
const GRAPHICLESS_DRAWING = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="Chart 2"/></wp:inline></w:drawing>`;

type BodyItem =
  | { kind: "paragraph"; text: string }
  | { kind: "picture" }
  | { kind: "graphicless" }
  | { kind: "picture-in-text"; text: string }
  | { kind: "graphicless-in-text"; text: string }
  | { kind: "table"; rows: number; cellText: string };

const tableXml = (rows: number, cellText: string): string => {
  const cells = (row: number): string =>
    [0, 1]
      .map(
        (column) =>
          `<w:tc><w:tcPr><w:tcW w:w="1870" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">r${String(row)}c${String(column)} ${cellText}</w:t></w:r></w:p></w:tc>`,
      )
      .join("");
  const body = Array.from({ length: rows }, (_unused, row) => `<w:tr>${cells(row)}</w:tr>`).join(
    "",
  );
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1870"/><w:gridCol w:w="1870"/></w:tblGrid>${body}</w:tbl>`;
};

const bodyItemXml = (item: BodyItem): string => {
  switch (item.kind) {
    case "paragraph":
      return `<w:p><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r></w:p>`;
    case "picture":
      return `<w:p><w:r>${PICTURE}</w:r></w:p>`;
    case "graphicless":
      return `<w:p><w:r>${GRAPHICLESS_DRAWING}</w:r></w:p>`;
    case "picture-in-text":
      return `<w:p><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r><w:r>${PICTURE}</w:r><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r></w:p>`;
    case "graphicless-in-text":
      return `<w:p><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r><w:r>${GRAPHICLESS_DRAWING}</w:r><w:r><w:t xml:space="preserve">${item.text}</w:t></w:r></w:p>`;
    case "table":
      return tableXml(item.rows, item.cellText);
    default:
      return item satisfies never;
  }
};

type PackageSpec = { items: readonly BodyItem[]; styleDefinitions: boolean };

const buildPackage = async ({ items, styleDefinitions }: PackageSpec): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  const stylesOverride = styleDefinitions
    ? `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`
    : "";
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${stylesOverride}</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const stylesRelationship = styleDefinitions
    ? `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/>`
    : "";
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECL}<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">${stylesRelationship}<Relationship Id="${IMAGE_RID}" Type="${OFFICE_RELATIONSHIPS}/image" Target="media/image1.png"/></Relationships>`,
  );
  if (styleDefinitions) {
    zip.file(
      "word/styles.xml",
      `${XML_DECL}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    );
  }
  zip.file("word/media/image1.png", PNG_BYTES);
  zip.file(
    "word/document.xml",
    `${XML_DECL}<w:document ${NAMESPACES}><w:body>${items
      .map(bodyItemXml)
      .join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const sentence = fc.constantFrom(
  "The parties agree to the following terms.",
  "This clause survives termination.",
  "Notice is given in writing.",
);

/**
 * Row counts on both sides of the content-structure profile's block cap, and
 * cell texts on both sides of its text cap.
 *
 * Each profile cap blanked the exact signature it was computed from, so a table
 * one block past 128 blocks, or one code unit past the retained text budget,
 * carried no evidence that it was itself: the engine reported the only table in
 * a document deleted and re-inserted when it compared the document with itself.
 */
const tableArbitrary: fc.Arbitrary<BodyItem> = fc.oneof(
  fc.record({
    kind: fc.constant("table" as const),
    rows: fc.constantFrom(1, 4, 64, 65),
    cellText: fc.constantFrom("", "lorem ipsum dolor"),
  }),
  fc.record({
    kind: fc.constant("table" as const),
    rows: fc.constantFrom(1, 3),
    cellText: fc.constantFrom("short", "padding ".repeat(1_100)),
  }),
);

const bodyItemArbitrary: fc.Arbitrary<BodyItem> = fc.oneof(
  sentence.map<BodyItem>((text) => ({ kind: "paragraph", text })),
  fc.constant<BodyItem>({ kind: "picture" }),
  fc.constant<BodyItem>({ kind: "graphicless" }),
  sentence.map<BodyItem>((text) => ({ kind: "picture-in-text", text })),
  sentence.map<BodyItem>((text) => ({ kind: "graphicless-in-text", text })),
  tableArbitrary,
);

const packageArbitrary: fc.Arbitrary<PackageSpec> = fc.record({
  items: fc.array(bodyItemArbitrary, { minLength: 1, maxLength: 6 }),
  styleDefinitions: fc.boolean(),
});

describe("compare(x, x) (property)", () => {
  test(
    "reports no change and verifies, for every package the engine parses",
    async () => {
      await fc.assert(
        fc.asyncProperty(packageArbitrary, async (spec) => {
          const buffer = await buildPackage(spec);
          // Two independent copies: a single buffer passed twice would let
          // either side's handling alias the other's bytes.
          const compared = await compareDocx(buffer.slice(0), buffer.slice(0), COMPARE_OPTIONS);
          if (compared.isErr()) {
            throw new Error(`compare(x, x) failed: ${compared.error.message}`);
          }
          expect(compared.value.changes).toEqual([]);
          // `verification` is where the other two clauses are certified:
          // `accept-reproduces-target` and `reject-reproduces-base` both hold
          // exactly when it is `verified`, with no failure to name.
          expect(compared.value.verification).toEqual({ status: "verified" });
          expect(compared.value.compatibility.status).toBe("standard-ooxml");
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(120_000),
  );

  // The law says identical input reports nothing. It must not have been bought
  // by making everything past a profile cap compare equal: the digest that
  // carries identity across the cap has to separate different content too.
  test("still reports an edit inside a table past the profile cap", async () => {
    const closing = { kind: "paragraph", text: "Executed as a deed." } as const;
    const base = await buildPackage({
      items: [{ kind: "table", rows: 65, cellText: "as originally drafted" }, closing],
      styleDefinitions: true,
    });
    const revised = await buildPackage({
      items: [{ kind: "table", rows: 65, cellText: "as subsequently amended" }, closing],
      styleDefinitions: true,
    });

    const compared = await compareDocx(base, revised, COMPARE_OPTIONS);

    if (compared.isErr()) {
      throw new Error(`compare reported an error: ${compared.error.message}`);
    }
    expect(compared.value.changes.length).toBeGreaterThan(0);
  });
});
