import { describe, expect, test } from "bun:test";

import { serializeTableCellFormatting } from "../serializer/tableSerializer";
import { parseStyles } from "../styleParser";
import { parseTableCellProperties } from "../tableParser";
import type { XmlElement } from "../xmlParser";
import { parseXmlDocument } from "../xmlParser";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe("diagonal table cell borders", () => {
  test("parses and serializes both cell diagonal directions", () => {
    const tcPr = parseXmlDocument(
      `<w:tcPr ${W_NS}><w:tcBorders><w:tl2br w:val="dashed" w:sz="12" w:color="123456"/><w:tr2bl w:val="double" w:sz="18" w:color="654321"/></w:tcBorders></w:tcPr>`,
    ) as XmlElement | null;

    const formatting = parseTableCellProperties(tcPr);

    expect(formatting?.borders?.topLeftToBottomRight).toMatchObject({
      color: { rgb: "123456" },
      size: 12,
      style: "dashed",
    });
    expect(formatting?.borders?.topRightToBottomLeft).toMatchObject({
      color: { rgb: "654321" },
      size: 18,
      style: "double",
    });

    const serialized = serializeTableCellFormatting(formatting);
    expect(serialized).toContain('<w:tl2br w:val="dashed" w:sz="12" w:color="123456"/>');
    expect(serialized).toContain('<w:tr2bl w:val="double" w:sz="18" w:color="654321"/>');
  });

  test("preserves diagonals declared by a table style cell property", () => {
    const styles = parseStyles(
      `<w:styles ${W_NS}>
        <w:style w:type="table" w:styleId="DiagonalCells">
          <w:name w:val="Diagonal cells"/>
          <w:tcPr>
            <w:tcBorders><w:tl2br w:val="single" w:sz="8"/></w:tcBorders>
          </w:tcPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("DiagonalCells")?.tcPr?.borders?.topLeftToBottomRight).toMatchObject({
      size: 8,
      style: "single",
    });
  });
});
