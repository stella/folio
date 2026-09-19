/**
 * `w:shd`'s `w:fill` and `w:color` are `ST_HexColor`: six hex digits or the
 * reserved token `auto`. `auto` paints nothing, but it is not the same as an
 * absent attribute, and three of the four `w:shd` parsers dropped it, so the
 * sentinel survived only on runs. Saving a cell that cancels its table style's
 * fill with `w:fill="auto"` wrote the cancellation away.
 *
 * The cascade itself held even while the sentinel was dropped, because
 * `CT_Shd/@w:val` is required: a schema-valid `w:shd` always parses to a
 * non-empty shading, which takes the direct branch of the cell cascade. The
 * first property pins that so the collapsed parser cannot regress it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { expectTableCellAttrs } from "../prosemirror/attrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { parseParagraphProperties } from "./paragraphParser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";
import { parseRun } from "./runParser";
import { parseStyles } from "./styleParser";
import { parseTableCellProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** `auto`, a valid hex colour, and a value that is neither. */
const FILL_VALUES = ["auto", "E0E0E0", "notacolour"] as const;

type FillValue = (typeof FILL_VALUES)[number];

const shd = (fill: FillValue): string => `<w:shd w:val="clear" w:fill="${fill}"/>`;

const expectedFill = (fill: FillValue) => {
  switch (fill) {
    case "auto":
      return { auto: true };
    case "E0E0E0":
      return { rgb: fill };
    default:
      return undefined;
  }
};

const parseOne = (xml: string) => {
  const element = parseXmlDocument(xml);
  if (!element) {
    throw new Error("fixture did not parse");
  }
  return element;
};

const runShading = (fill: FillValue) =>
  parseRun(parseOne(`<w:r ${WORD_NAMESPACE}><w:rPr>${shd(fill)}</w:rPr></w:r>`), null, null)
    .formatting?.shading;

const paragraphShading = (fill: FillValue) =>
  parseParagraphProperties(parseOne(`<w:pPr ${WORD_NAMESPACE}>${shd(fill)}</w:pPr>`), null)
    ?.shading;

const styleShading = (fill: FillValue) =>
  parseStyles(
    `<w:styles ${WORD_NAMESPACE}>
       <w:style w:type="paragraph" w:styleId="Shaded">
         <w:name w:val="Shaded"/>
         <w:pPr>${shd(fill)}</w:pPr>
       </w:style>
     </w:styles>`,
    null,
  ).get("Shaded")?.pPr?.shading;

const cellShading = (fill: FillValue) =>
  parseTableCellProperties(parseOne(`<w:tcPr ${WORD_NAMESPACE}>${shd(fill)}</w:tcPr>`))?.shading;

const styledTableFixture = (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
</Relationships>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles ${WORD_NAMESPACE}>
  <w:style w:type="table" w:styleId="FilledTable">
    <w:name w:val="Filled Table"/>
    <w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr>
  </w:style>
</w:styles>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document ${WORD_NAMESPACE}>
  <w:body>
    <w:p><w:r><w:t>Table style fill</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="FilledTable"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:shd w:val="clear" w:fill="auto"/></w:tcPr>
          <w:p><w:r><w:t>Cancels the style fill</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Takes the style fill</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: "arraybuffer" });
};

describe("w:shd auto", () => {
  test("an explicit auto fill on a cell cancels the table-style fill", async () => {
    const original = await parseDocx(await styledTableFixture());
    const row = toProseDoc(original).child(1).firstChild;
    if (!row) {
      throw new Error("Expected a table row");
    }

    expect(expectTableCellAttrs(row.child(0)).backgroundColor).toBeUndefined();
    expect(expectTableCellAttrs(row.child(1)).backgroundColor).toBe("FF0000");
  }, 15_000);

  test("the sentinel survives a save", async () => {
    const original = await parseDocx(await styledTableFixture());
    const rebuilt = fromProseDoc(toProseDoc(original), original);
    const repacked = await repackDocx(rebuilt, { updateModifiedDate: false });
    const documentXml = await (
      await JSZip.loadAsync(repacked)
    )
      .file("word/document.xml")
      ?.async("string");

    expect(documentXml).toContain('w:fill="auto"');
  }, 15_000);

  test(
    "every tier reads a fill the same way",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...FILL_VALUES), (fill) => {
          const tiers = [
            runShading(fill),
            paragraphShading(fill),
            styleShading(fill),
            cellShading(fill),
          ];

          for (const tier of tiers) {
            expect(tier?.fill).toEqual(expectedFill(fill));
          }
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );
});
