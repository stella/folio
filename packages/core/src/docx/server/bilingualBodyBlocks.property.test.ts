import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { BlockContent } from "../../types/document";
import { parseDocx } from "../parser";
import { createEmptyDocx, createDocx } from "../rezip";
import {
  createBilingualDocument,
  readBilingualDocument,
  type BilingualTableLayout,
} from "./createBilingualDocument";

const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const CHUNK_PART = "word/imported.html";
const CHUNK_CONTENT = "<html><body><p>Imported clause</p></body></html>";

const paragraph = (id: number): string =>
  `<w:p w14:paraId="${id.toString(16).padStart(8, "0").toUpperCase()}"><w:r><w:t>Clause ${id}</w:t></w:r></w:p>`;

const table = (id: number): string =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${paragraph(id)}</w:tc></w:tr></w:tbl>`;

const OPAQUE_UNITS = {
  permission: (id: number) =>
    `<w:permStart w:id="${id}" w:edGrp="everyone"/>${paragraph(id)}<w:permEnd w:id="${id}"/>`,
  bookmark: (id: number) =>
    `<w:bookmarkStart w:id="${id}" w:name="body_${id}"/>${paragraph(id)}<w:bookmarkEnd w:id="${id}"/>`,
  altChunk: () => '<w:altChunk r:id="rIdChunk"/>',
  math: (id: number) =>
    `<m:oMathPara><m:oMath><m:r><m:t>x${id}</m:t></m:r></m:oMath></m:oMathPara>`,
} as const;

const UNITS = { ...OPAQUE_UNITS, paragraph, table } as const;
const PLACEMENTS = {
  body: (xml: string) => xml,
  sdt: (xml: string) =>
    `<w:sdt><w:sdtPr><w:tag w:val="source"/></w:sdtPr><w:sdtContent>${xml}</w:sdtContent></w:sdt>`,
  customXml: (xml: string) => `<w:customXml w:element="clause">${xml}</w:customXml>`,
} as const;

const buildPackage = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}" xmlns:m="${M_NAMESPACE}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  expect(rels).toBeDefined();
  zip.file(
    "word/_rels/document.xml.rels",
    (rels ?? "").replace(
      "</Relationships>",
      `<Relationship Id="rIdChunk" Type="${R_NAMESPACE}/aFChunk" Target="imported.html"/></Relationships>`,
    ),
  );
  const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
  expect(contentTypes).toBeDefined();
  zip.file(
    "[Content_Types].xml",
    (contentTypes ?? "").replace(
      "</Types>",
      '<Override PartName="/word/imported.html" ContentType="text/html"/></Types>',
    ),
  );
  zip.file(CHUNK_PART, CHUNK_CONTENT);
  return zip.generateAsync({ type: "arraybuffer" });
};

const paragraphIds = (blocks: readonly BlockContent[]): Set<string> => {
  const ids = new Set<string>();
  const visit = (block: BlockContent): void => {
    switch (block.type) {
      case "paragraph":
        if (block.paraId !== undefined) ids.add(block.paraId);
        return;
      case "table":
        for (const row of block.rows) for (const cell of row.cells) cell.content.forEach(visit);
        return;
      case "blockSdt":
      case "blockCustomXml":
        block.content.forEach(visit);
        return;
      case "bookmarkStart":
      case "bookmarkEnd":
      case "preservedBlock":
        return;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  };
  blocks.forEach(visit);
  return ids;
};

// Capture every carried block, including duplicate opaque markup. Equality
// checks both its multiplicity and its order across nested source wrappers.
const carriedBlocks = (blocks: readonly BlockContent[]): string[] => {
  const carried: string[] = [];
  const visit = (block: BlockContent): void => {
    switch (block.type) {
      case "preservedBlock":
        carried.push(block.xml);
        return;
      case "bookmarkStart":
      case "bookmarkEnd":
        carried.push(JSON.stringify(block));
        return;
      case "blockSdt":
      case "blockCustomXml":
        block.content.forEach(visit);
        return;
      case "table":
        for (const row of block.rows) for (const cell of row.cells) cell.content.forEach(visit);
        return;
      case "paragraph":
        return;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  };
  blocks.forEach(visit);
  return carried;
};

const roundTrip = async (body: string, tableLayout: BilingualTableLayout) => {
  const source = await parseDocx(await buildPackage(body), { preloadFonts: false });
  const options = {
    targetStyleSuffix: "en",
    tableLayout,
    editableParagraphIds: paragraphIds(source.package.document.content),
  };
  const result = createBilingualDocument(source, options);
  expect(createBilingualDocument(source, options).rows).toEqual(result.rows);
  expect(carriedBlocks(result.document.package.document.content)).toEqual(
    carriedBlocks(source.package.document.content),
  );
  const saved = await createDocx(result.document);
  const reparsed = await parseDocx(saved, { preloadFonts: false });
  expect(carriedBlocks(reparsed.package.document.content)).toEqual(
    carriedBlocks(source.package.document.content),
  );
  expect(readBilingualDocument(reparsed, paragraphIds(reparsed.package.document.content))).toEqual(
    result.rows,
  );
  return { source, result, reparsed, zip: await JSZip.loadAsync(saved) };
};

describe("bilingual body block preservation", () => {
  test("parsed packages retain range pairs, display math and altChunk targets inside and outside SDTs", async () => {
    let id = 100;
    const opaqueBody = Object.values(PLACEMENTS)
      .map((place) =>
        place(
          Object.values(OPAQUE_UNITS)
            .map((unit) => unit(id++))
            .join(""),
        ),
      )
      .join("");
    const body =
      opaqueBody +
      `<w:bookmarkStart w:id="600" w:name="body_to_inline"/>${paragraph(601).replace("</w:p>", '<w:bookmarkEnd w:id="600"/></w:p>')}` +
      `${paragraph(603).replace("<w:r>", '<w:bookmarkStart w:id="602" w:name="inline_to_body"/><w:r>')}<w:bookmarkEnd w:id="602"/>`;
    for (const layout of ["inline", "stacked"] as const) {
      const { source, zip } = await roundTrip(body, layout);
      expect(carriedBlocks(source.package.document.content)).toHaveLength(20);
      const xml = await zip.file("word/document.xml")?.async("text");
      for (const element of [
        "permStart",
        "permEnd",
        "bookmarkStart",
        "bookmarkEnd",
        "altChunk",
        "oMathPara",
      ]) {
        expect(xml).toContain(element);
      }
      for (const bookmarkId of [600, 602]) {
        expect(
          xml?.match(new RegExp(`<w:bookmarkStart\\b[^>]*w:id="${bookmarkId}"`, "g")),
        ).toHaveLength(1);
        expect(
          xml?.match(new RegExp(`<w:bookmarkEnd\\b[^>]*w:id="${bookmarkId}"`, "g")),
        ).toHaveLength(1);
      }
      expect(await zip.file(CHUNK_PART)?.async("text")).toBe(CHUNK_CONTENT);
      expect(await zip.file("word/_rels/document.xml.rels")?.async("text")).toContain(
        'Target="imported.html"',
      );
      expect(await zip.file("[Content_Types].xml")?.async("text")).toContain(
        'PartName="/word/imported.html"',
      );
    }
  });

  test(
    "every non-stamped block survives random mixed body sequences in source order",
    async () => {
      await assertProperty(
        fc.asyncProperty(
          fc.array(
            fc.record({
              unit: fc.constantFrom(...Object.values(UNITS)),
              place: fc.constantFrom(...Object.values(PLACEMENTS)),
            }),
            { minLength: 1, maxLength: 16 },
          ),
          fc.constantFrom("inline", "stacked"),
          async (sequence, layout) => {
            const body = sequence
              .map(({ unit, place }, index) => place(unit(index + 1000)))
              .join("");
            await roundTrip(body, layout);
          },
        ),
        { id: "bilingual carries every non-stamped body block in source order", numRuns: 30 },
      );
    },
    propertyTestTimeout(30_000),
  );
});
