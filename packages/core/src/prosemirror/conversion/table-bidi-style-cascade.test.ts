import { describe, expect, test } from "bun:test";

import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import type { Document, TableFormatting } from "../../types/document";
import { toProseDoc } from "./toProseDoc";

const tableDocument = (formatting: TableFormatting): Document => ({
  package: {
    styles: {
      styles: [
        {
          styleId: "RtlTable",
          type: "table",
          tblPr: { bidi: true },
        },
      ],
    },
    document: {
      content: [
        {
          type: "table",
          formatting,
          rows: [
            {
              cells: [{ content: [{ type: "paragraph", content: [] }] }],
            },
          ],
        },
      ],
    },
  },
});

const tableAttrs = (document: Document) =>
  toProseDoc(document, { styles: document.package.styles }).firstChild?.attrs;

const tableBidi = (document: Document) => {
  const proseDoc = toProseDoc(document, { styles: document.package.styles });
  const table = toFlowBlocks(proseDoc).find((block) => block.kind === "table");
  return table?.kind === "table" ? table.bidi : undefined;
};

describe("table bidiVisual style cascade", () => {
  test("resolves RTL from the table style without copying it into direct formatting", () => {
    const attrs = tableAttrs(tableDocument({ styleId: "RtlTable" }));

    expect(attrs?._resolvedBidi).toBe(true);
    expect(attrs?._originalFormatting).toEqual({ styleId: "RtlTable" });
    expect(tableBidi(tableDocument({ styleId: "RtlTable" }))).toBe(true);
  });

  test("direct explicit LTR overrides an RTL table style", () => {
    const attrs = tableAttrs(tableDocument({ styleId: "RtlTable", bidi: false }));

    expect(attrs?._resolvedBidi).toBe(false);
    expect(attrs?._originalFormatting).toEqual({ styleId: "RtlTable", bidi: false });
    expect(tableBidi(tableDocument({ styleId: "RtlTable", bidi: false }))).toBeUndefined();
  });
});
