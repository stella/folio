/**
 * A table width keeps its unit through a save, whichever way it was spelled.
 *
 * `CT_TblWidth` carries `w:w` typed `ST_MeasurementOrPercent` and an *optional*
 * `w:type` the schema gives no default. Folio read a missing `w:type` as
 * `dxa`, so `<w:tblW w:w="50%"/>` came back as 50 twips and was written back
 * that way: a full-width table became a hairline, silently, in a file Word
 * lays out correctly.
 *
 * Both spellings of a percentage are generated, because they are the same
 * value and folio has to agree with itself about that: `50%` and the `2500`
 * its slot counts in are one width.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** `w:type`, plus the absence the schema allows and folio used to read as `dxa`. */
const DECLARED_TYPES = ["absent", "auto", "dxa", "nil", "pct"] as const;

type DeclaredType = (typeof DECLARED_TYPES)[number];

/** How the number is written: as itself, or as the percentage it reads as. */
const SPELLINGS = ["number", "percentSign"] as const;

type Spelling = (typeof SPELLINGS)[number];

/** 50ths of a percent, the unit `ST_MeasurementOrPercent` counts in. */
const FIFTIETHS_PER_PERCENT = 50;

type Width = { declared: DeclaredType; spelling: Spelling; percent: number };

const attributesFor = ({ declared, spelling, percent }: Width): string => {
  const value =
    spelling === "percentSign" ? `${percent}%` : String(percent * FIFTIETHS_PER_PERCENT);
  const type = declared === "absent" ? "" : ` w:type="${declared}"`;
  return `w:w="${value}"${type}`;
};

/**
 * The width folio must end up with.
 *
 * A `%` spelling is a percentage whatever `w:type` says, because no number of
 * twips is written that way; `auto` and `nil` do not read `w:w` as a width at
 * all, so the declared type stands.
 */
const expectedFor = ({ declared, spelling, percent }: Width): { value: number; type: string } => {
  if (declared === "auto" || declared === "nil") {
    return {
      value: spelling === "percentSign" ? percent : percent * FIFTIETHS_PER_PERCENT,
      type: declared,
    };
  }
  if (spelling === "percentSign") {
    return { value: percent * FIFTIETHS_PER_PERCENT, type: "pct" };
  }
  return { value: percent * FIFTIETHS_PER_PERCENT, type: declared === "absent" ? "dxa" : declared };
};

const bodyFor = (width: Width): string =>
  `<w:tbl><w:tblPr><w:tblW ${attributesFor(width)}/></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW ${attributesFor(width)}/></w:tcPr>` +
  `<w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

type ReadWidths = { table: unknown; cell: unknown };

const widthsOf = (document: Awaited<ReturnType<typeof parseDocx>>): ReadWidths => {
  const table = document.package.document.content.find((block) => block.type === "table");
  if (table?.type !== "table") {
    throw new Error("the fixture lost its table");
  }
  return {
    table: table.formatting?.width,
    cell: table.rows.at(0)?.cells.at(0)?.formatting?.width,
  };
};

/** Parse, force a rebuild rather than a replay, and parse what was written. */
const throughSave = async (body: string): Promise<{ read: ReadWidths; resaved: ReadWidths }> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const read = widthsOf(parsed);
  const reopened = await parseDocx(await repackDocx(parsed, { updateModifiedDate: false }), {
    preloadFonts: false,
  });
  return { read, resaved: widthsOf(reopened) };
};

describe("a table width survives every spelling of its own unit", () => {
  test("every declared type and both spellings keep one effective width", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...DECLARED_TYPES),
        fc.constantFrom(...SPELLINGS),
        fc.integer({ min: 1, max: 100 }),
        async (declared, spelling, percent) => {
          const width = { declared, spelling, percent };
          const expected = expectedFor(width);
          const { read, resaved } = await throughSave(bodyFor(width));

          // The reader agrees with the spelling, and the save does not move it.
          expect({ declared, spelling, ...read }).toEqual({
            declared,
            spelling,
            table: expected,
            cell: expected,
          });
          expect({ declared, spelling, ...resaved }).toEqual({
            declared,
            spelling,
            table: expected,
            cell: expected,
          });
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("a percentage with no declared type never becomes twips", async () => {
    const { read, resaved } = await throughSave(
      bodyFor({ declared: "absent", spelling: "percentSign", percent: 50 }),
    );

    // 50 twips is 0.9mm. The bug wrote that for a table half the page wide.
    expect(read).toEqual({
      table: { value: 2500, type: "pct" },
      cell: { value: 2500, type: "pct" },
    });
    expect(resaved).toEqual(read);
  });
});
