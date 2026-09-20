import { describe, expect, test } from "bun:test";

import {
  serializeTableCellFormatting,
  serializeTableFormatting,
  serializeTableRowFormatting,
} from "./serializer/tableSerializer";
import { parseBlockContent } from "./blockContentParser";
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
  const table = parseTable(root, null, null, null, null, new Map());
  if (!table) {
    throw new Error("Expected a table with at least one row");
  }
  return table;
}

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * A property set as if it had been built rather than parsed.
 *
 * A parsed one carries the element it came from, and the serializer writes
 * that back verbatim; dropping it is how these cases reach the branch that
 * rebuilds the element from the typed values.
 */
const rebuilt = <TFormatting extends { sourceXml?: string }>(
  formatting: TFormatting | undefined,
): TFormatting | undefined => {
  if (!formatting) {
    return formatting;
  }
  const { sourceXml: _source, ...rest } = formatting;
  return rest as TFormatting;
};

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

describe("rowless tables", () => {
  test("omits a placeholder table instead of creating an invalid model", () => {
    const root = parseXmlDocument(`<w:body ${NS}>
      <w:p><w:r><w:t>before</w:t></w:r></w:p>
      <w:tbl><w:tblPr/></w:tbl>
      <w:p><w:r><w:t>after</w:t></w:r></w:p>
    </w:body>`) as XmlElement;

    expect(
      parseBlockContent(root, null, null, null, null, new Map()).map(({ type }) => type),
    ).toEqual(["paragraph", "paragraph"]);
  });
});

describe("table cell marker visibility", () => {
  test("writes the package's one off spelling for an explicit false override", () => {
    const root = parseXmlDocument(`<w:tcPr ${NS}><w:hideMark w:val="false"/></w:tcPr>`);
    const formatting = parseTableCellProperties(root);

    expect(formatting?.hideMark).toBe(false);
    // The parsed element goes back exactly as it arrived, `w:val="false"` and
    // all; the rebuilt one writes `0`, as every other on/off element does.
    expect(serializeTableCellFormatting(formatting)).toContain('<w:hideMark w:val="false"/>');
    expect(serializeTableCellFormatting(rebuilt(formatting))).toContain('<w:hideMark w:val="0"/>');
  });
});

describe("table cell grid span cap", () => {
  test("clamps a hostile gridSpan to the practical column cap", () => {
    const root = parseXmlDocument(`<w:tcPr ${NS}><w:gridSpan w:val="2000000000"/></w:tcPr>`);
    const formatting = parseTableCellProperties(root);

    expect(formatting?.gridSpan).toBe(63);
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

describe("table properties", () => {
  test("writes w:tblPr children in the order CT_TblPrBase declares", () => {
    const xml = serializeTableFormatting({
      styleId: "TableGrid",
      overlap: "never",
      bidi: false,
      width: { value: 8000, type: "dxa" },
      justification: "center",
      indent: { value: 120, type: "dxa" },
      borders: { top: { style: "single", size: 4, space: 0, color: { rgb: "000000" } } },
      shading: { fill: { rgb: "D9D9D9" }, pattern: "clear" },
      layout: "fixed",
      cellMargins: { top: { value: 80, type: "dxa" } },
      look: { firstRow: true },
    });

    // CT_TblPrBase is a sequence, so a consumer validating the part refuses
    // any other order. Pinned because the elements are written from separate
    // branches that read the same object.
    const order = [...xml.matchAll(/<w:([A-Za-z]+)/gu)].map(([, name]) => name);
    expect(order).toEqual([
      "tblPr",
      "tblStyle",
      "tblOverlap",
      "bidiVisual",
      "tblW",
      "jc",
      "tblInd",
      "tblBorders",
      "top",
      "shd",
      "tblLayout",
      "tblCellMar",
      "top",
      "tblLook",
    ]);
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
    // Parsed, the element is written back as it arrived — the fixture's own
    // indentation included. Rebuilt, the children come out in the order
    // `CT_TrPrBase` declares them, which pairs the two counts before the two
    // widths rather than pairing each count with its own width.
    expect(serializeTableRowFormatting(formatting)).toContain('<w:gridBefore w:val="2"/>');
    expect(serializeTableRowFormatting(rebuilt(formatting))).toContain(
      '<w:gridBefore w:val="2"/><w:gridAfter w:val="1"/><w:wBefore w:w="900" w:type="dxa"/><w:wAfter w:w="450" w:type="dxa"/>',
    );
  });
});

describe("table row height", () => {
  test("keeps a non-positive height the reader takes no value from", () => {
    const root = parseXmlDocument(
      `<w:trPr ${NS}><w:trHeight w:val="0" w:hRule="atLeast"/></w:trPr>`,
    );
    const formatting = parseTableRowProperties(root);

    // A height of zero is a value the reader refuses, not an element folio has
    // never heard of, so it goes to the sink rather than off the end of the
    // walk: the model states no height and the rule goes back with it.
    expect(formatting?.height).toBeUndefined();
    expect(formatting?.heightRule).toBeUndefined();
    expect(serializeTableRowFormatting(rebuilt(formatting))).toBe(
      '<w:trPr><w:trHeight w:val="0" w:hRule="atLeast"/></w:trPr>',
    );
  });
});

describe("table row conditional formatting", () => {
  test("round-trips conditional table-style flags", () => {
    const root = parseXmlDocument(
      `<w:trPr ${NS}><w:cnfStyle w:val="100000100000"/></w:trPr>`,
    ) as XmlElement;
    const formatting = parseTableRowProperties(root);

    expect(rebuilt(formatting)).toEqual({
      conditionalFormat: {
        firstRow: true,
        oddHBand: true,
      },
    });

    // Both ways round: the element as parsed, and the element the serializer
    // builds when it has no capture to write back.
    for (const serialized of [
      serializeTableRowFormatting(formatting),
      serializeTableRowFormatting(rebuilt(formatting)),
    ]) {
      expect(rebuilt(parseTableRowProperties(parseXmlDocument(serialized)))).toEqual(
        rebuilt(formatting),
      );
    }
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
  test("keeps a w:val outside ST_Border verbatim", () => {
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
      // `dashDotDot` is not an `ST_Border` member (`dotDotDash` is), so it
      // reaches the model as the token the file wrote, not as a style.
      style: { kind: "unrecognised", raw: "dashDotDot" },
    });
  });
});

describe("table bookmark placement", () => {
  test("a cell-level marker is a block of the cell and a row-level one rides the row", () => {
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

    const row = table.rows.at(0);
    const firstCell = row?.cells.at(0);
    // The cell's own range brackets the cell's paragraph, as the source wrote it.
    expect(firstCell?.content.map((block) => block.type)).toEqual([
      "bookmarkStart",
      "paragraph",
      "bookmarkEnd",
    ]);
    expect(firstCell?.content.at(0)).toMatchObject({ type: "bookmarkStart", id: 1 });
    expect(firstCell?.content.at(-1)).toMatchObject({ type: "bookmarkEnd", id: 1 });

    // The row-level end closes after the first cell, not inside it: the range
    // still covers the cell rather than part of its text.
    expect(row?.bookmarks).toEqual([{ index: 1, marker: { type: "bookmarkEnd", id: 2 } }]);
  });

  test("a bookmark that selects whole rows keeps both halves on the row", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:tr>
        <w:bookmarkStart w:id="7" w:name="wholeRow"/>
        <w:tc><w:p><w:r><w:t>First</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:tc>
        <w:bookmarkEnd w:id="7"/>
      </w:tr>
    </w:tbl>`);

    expect(table.rows.at(0)?.bookmarks).toEqual([
      { index: 0, marker: { type: "bookmarkStart", id: 7, name: "wholeRow" } },
      { index: 2, marker: { type: "bookmarkEnd", id: 7 } },
    ]);
  });

  test("a bookmark that selects a whole table keeps both halves on the table", () => {
    const table = parseTableXml(`<w:tbl ${NS}>
      <w:bookmarkStart w:id="9" w:name="wholeTable"/>
      <w:tr><w:tc><w:p><w:r><w:t>First</w:t></w:r></w:p></w:tc></w:tr>
      <w:bookmarkEnd w:id="9"/>
    </w:tbl>`);

    expect(table.bookmarks).toEqual([
      { index: 0, marker: { type: "bookmarkStart", id: 9, name: "wholeTable" } },
      { index: 1, marker: { type: "bookmarkEnd", id: 9 } },
    ]);
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
