import { describe, expect, test } from "bun:test";

import { parseStyleDefinitions } from "../../docx/styleParser";
import type { Document, Table } from "../../types/document";
import { toProseDoc } from "./toProseDoc";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const STYLES_XML = `<w:styles ${W}>
  <w:style w:type="table" w:default="1" w:styleId="PlainTable">
    <w:name w:val="Plain Table"/>
    <w:tblPr>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="100" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="100" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="Wide">
    <w:name w:val="Wide"/>
    <w:tblPr>
      <w:tblCellMar>
        <w:left w:w="400" w:type="dxa"/>
        <w:right w:w="400" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>`;

function tableWithStyle(styleId: string | undefined): Table {
  return {
    type: "table",
    ...(styleId ? { formatting: { styleId } } : {}),
    rows: [
      {
        type: "tableRow",
        cells: [{ type: "tableCell", content: [{ type: "paragraph", content: [] }] }],
      },
    ],
  };
}

function tableAttrs(styleId: string | undefined): Record<string, unknown> {
  const styles = parseStyleDefinitions(STYLES_XML, null);
  const document: Document = {
    package: { styles, document: { content: [tableWithStyle(styleId)] } },
  };
  return (toProseDoc(document, { styles }).firstChild?.attrs ?? {}) as Record<string, unknown>;
}

describe("w:tblStyle fallback to the default table style", () => {
  test("a table without w:tblStyle takes the default table style's cell margins", () => {
    expect(tableAttrs(undefined)["_resolvedCellMargins"]).toEqual({
      top: 0,
      left: 100,
      bottom: 0,
      right: 100,
    });
  });

  test("a w:tblStyle naming an undefined style takes the default table style", () => {
    const attrs = tableAttrs("NoSuchStyle");
    expect(attrs["_resolvedCellMargins"]).toEqual({ top: 0, left: 100, bottom: 0, right: 100 });
    // The authored reference is kept for the save.
    expect(attrs["styleId"]).toBe("NoSuchStyle");
  });

  test("a w:tblStyle naming a defined style does not consult the default", () => {
    expect(tableAttrs("Wide")["_resolvedCellMargins"]).toEqual({ left: 400, right: 400 });
  });
});
