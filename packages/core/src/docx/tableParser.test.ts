import { describe, expect, test } from "bun:test";

import {
  serializeTableCellFormatting,
  serializeTableRowFormatting,
} from "./serializer/tableSerializer";
import {
  parseTable,
  parseTableCellProperties,
  parseTableMeasurement,
  parseTableRowProperties,
} from "./tableParser";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

function parseTableXml(xml: string) {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse table XML");
  }
  return parseTable(root, null, null, null, null, new Map());
}

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe("parseTableMeasurement", () => {
  const tblW = (w: string, type = "pct") => {
    const root = parseXmlDocument(`<w:tblW ${NS} w:w="${w}" w:type="${type}"/>`) as XmlElement;
    return parseTableMeasurement(root);
  };

  test("normalizes percent-suffixed pct widths to 50ths-of-percent", () => {
    expect(tblW("100%")).toEqual({ value: 5000, type: "pct" });
    expect(tblW("50%")).toEqual({ value: 2500, type: "pct" });
    expect(tblW(" 100% ")).toEqual({ value: 5000, type: "pct" });
  });

  test("keeps canonical pct integers unchanged", () => {
    expect(tblW("5000")).toEqual({ value: 5000, type: "pct" });
  });
});

describe("table cell marker visibility", () => {
  test("preserves an explicit false value so it can override a table style", () => {
    const root = parseXmlDocument(`<w:tcPr ${NS}><w:hideMark w:val="false"/></w:tcPr>`);
    const formatting = parseTableCellProperties(root);

    expect(formatting?.hideMark).toBe(false);
    expect(serializeTableCellFormatting(formatting)).toContain('<w:hideMark w:val="false"/>');
  });
});

describe("table cell merge revisions", () => {
  test("preserves original and applied vertical merge states", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:cellMerge w:id="17" w:author="Reviewer" w:vMerge="cont" w:vMergeOrig="rest"/>
          </w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>`);

    const change = table.rows.at(0)?.cells.at(0)?.structuralChange;
    expect(change).toEqual({
      type: "tableCellMerge",
      info: { id: 17, author: "Reviewer" },
      verticalMerge: "continue",
      verticalMergeOriginal: "rest",
    });
    expect(serializeTableCellFormatting(undefined, undefined, change)).toContain(
      '<w:cellMerge w:id="17" w:author="Reviewer" w:vMerge="cont" w:vMergeOrig="rest"/>',
    );
  });

  test("preserves omitted revision states", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:cellMerge w:id="18" w:author="Reviewer"/></w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>`);

    const change = table.rows.at(0)?.cells.at(0)?.structuralChange;
    expect(change).toEqual({
      type: "tableCellMerge",
      info: { id: 18, author: "Reviewer" },
    });
    expect(serializeTableCellFormatting(undefined, undefined, change)).toContain(
      '<w:cellMerge w:id="18" w:author="Reviewer"/>',
    );
  });
});

describe("table row grid offsets", () => {
  test("preserves omitted leading and trailing columns", () => {
    const root = parseXmlDocument(`<w:trPr ${NS}>
      <w:gridBefore w:val="2"/>
      <w:wBefore w:w="900" w:type="dxa"/>
      <w:gridAfter w:val="1"/>
      <w:wAfter w:w="450" w:type="dxa"/>
    </w:trPr>`) as XmlElement;

    const formatting = parseTableRowProperties(root);

    expect(formatting).toMatchObject({
      gridBefore: 2,
      widthBefore: { value: 900, type: "dxa" },
      gridAfter: 1,
      widthAfter: { value: 450, type: "dxa" },
    });
    expect(serializeTableRowFormatting(formatting)).toContain(
      '<w:gridBefore w:val="2"/><w:wBefore w:w="900" w:type="dxa"/><w:gridAfter w:val="1"/><w:wAfter w:w="450" w:type="dxa"/>',
    );
  });
});

describe("inferImplicitSingleCellRowSpans", () => {
  test("does not expand a vMerge continuation single-cell row", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>
      </w:tr>
    </w:tbl>`);

    expect(table.rows[1]?.cells[0]?.formatting?.gridSpan ?? 1).toBe(1);
  });

  test("keeps explicit gridSpan and expands full-width single-cell rows", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
    </w:tbl>`);

    expect(table.rows[1]?.cells[0]?.formatting?.gridSpan ?? 1).toBe(1);
    expect(table.rows[2]?.cells[0]?.formatting?.gridSpan).toBe(2);
    expect(table.rows[3]?.cells[0]?.formatting?.gridSpan).toBe(3);
  });

  test("does not expand a single-cell row with explicit grid offsets", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
    </w:tbl>`);

    expect(table.rows[1]?.cells[0]?.formatting?.gridSpan ?? 1).toBe(1);
  });

  test("does not expand a narrow single-cell row without span evidence", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>
      </w:tr>
    </w:tbl>`);

    expect(table.rows[1]?.cells[0]?.formatting?.gridSpan ?? 1).toBe(1);
  });
});

describe("table structured document tag wrappers", () => {
  test("extracts cells wrapped by row-level content controls", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tr>
        <w:sdt>
          <w:sdtPr><w:alias w:val="Cell control"/></w:sdtPr>
          <w:sdtContent>
            <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc>
            <w:tc><w:p/></w:tc>
          </w:sdtContent>
        </w:sdt>
      </w:tr>
    </w:tbl>`);

    expect(table.rows).toHaveLength(1);
    expect(table.rows[0]?.cells).toHaveLength(2);
    expect(table.rows[0]?.cells[0]?.formatting?.vMerge).toBe("restart");
  });

  test("extracts rows wrapped by table-level content controls", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:sdt>
        <w:sdtPr><w:alias w:val="Row control"/></w:sdtPr>
        <w:sdtContent>
          <w:tr><w:tc><w:p/></w:tc></w:tr>
          <w:tr><w:tc><w:p/></w:tc></w:tr>
        </w:sdtContent>
      </w:sdt>
    </w:tbl>`);

    expect(table.rows).toHaveLength(2);
    expect(table.rows[0]?.cells).toHaveLength(1);
    expect(table.rows[1]?.cells).toHaveLength(1);
  });
});

describe("table borders", () => {
  test("preserves unknown border styles for fallback rendering", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="dashDotDot" w:sz="8" w:color="00AAFF"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr><w:tc><w:p/></w:tc></w:tr>
    </w:tbl>`);

    expect(table.formatting?.borders?.top).toMatchObject({
      color: { rgb: "00AAFF" },
      size: 8,
      style: "dashDotDot",
    });
  });
});

describe("table bookmark placement", () => {
  test("attaches cell-level and row-level bookmark markers to adjacent paragraphs", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tr>
        <w:tc>
          <w:bookmarkStart w:id="1" w:name="cellRange"/>
          <w:p>
            <w:r><w:t>First</w:t></w:r>
            <w:bookmarkStart w:id="2" w:name="rowRange"/>
          </w:p>
          <w:bookmarkEnd w:id="1"/>
        </w:tc>
        <w:bookmarkEnd w:id="2"/>
        <w:tc><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`);

    const firstParagraph = table.rows.at(0)?.cells.at(0)?.content.at(0);
    expect(firstParagraph?.type).toBe("paragraph");
    if (!firstParagraph || firstParagraph.type !== "paragraph") {
      return;
    }

    expect(firstParagraph.content.map((content) => content.type)).toEqual([
      "bookmarkStart",
      "run",
      "bookmarkStart",
      "bookmarkEnd",
      "bookmarkEnd",
    ]);
    expect(firstParagraph.content.at(0)).toMatchObject({
      type: "bookmarkStart",
      id: 1,
    });
    expect(firstParagraph.content.at(-2)).toMatchObject({
      type: "bookmarkEnd",
      id: 1,
    });
    expect(firstParagraph.content.at(-1)).toMatchObject({
      type: "bookmarkEnd",
      id: 2,
    });
  });
});

describe("parseTable pct width", () => {
  test("parses a full-width pct table from a percent-suffixed tblW", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tblPr><w:tblW w:w="100%" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000" w:type="pct"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr>
    </w:tbl>`);

    expect(table.formatting?.width).toEqual({ value: 5000, type: "pct" });
  });
});
