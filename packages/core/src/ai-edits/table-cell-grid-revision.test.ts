import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Table, TableCell, TrackedChangeInfo } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const GRID_WIDTHS = [1200, 1800, 2400, 4680];
const GRID_SOURCE =
  '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1800"/><w:gridCol w:w="2400"/><w:gridCol w:w="4680"/></w:tblGrid>';
const UNCHANGED_GRID_SOURCE =
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="7000"/></w:tblGrid>';
const TABLE_SOURCE = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>';

const revision = (id: number): TrackedChangeInfo => ({
  id,
  author: "Reviewer",
  date: "2026-01-02T03:04:05Z",
});

const cell = (text: string, revisionId?: number): TableCell => ({
  type: "tableCell",
  ...(revisionId !== undefined && {
    structuralChange: { type: "tableCellDeletion", info: revision(revisionId) },
  }),
  content: [
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

const table = ({
  widths,
  gridSourceXml,
  cells,
}: {
  widths: number[];
  gridSourceXml: string;
  cells: TableCell[];
}): Table => ({
  type: "table",
  columnWidths: widths,
  formatting: { sourceXml: TABLE_SOURCE, gridSourceXml },
  rows: [{ type: "tableRow", cells }],
});

const buildSource = (): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [
          table({
            widths: GRID_WIDTHS,
            gridSourceXml: GRID_SOURCE,
            cells: [cell("A", 101), cell("B", 102), cell("C", 103), cell("Survivor")],
          }),
          table({
            widths: [3000, 7000],
            gridSourceXml: UNCHANGED_GRID_SOURCE,
            cells: [cell("Unaffected left"), cell("Unaffected right")],
          }),
        ],
      },
    },
  });
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!entry) {
    throw new Error("expected word/document.xml");
  }
  return entry.async("string");
};

const gridsFrom = (xml: string): string[] => xml.match(/<w:tblGrid>.*?<\/w:tblGrid>/gu) ?? [];

const tablesFrom = async (buffer: ArrayBuffer): Promise<Table[]> => {
  const parsed = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  return parsed.package.document.content.filter((block): block is Table => block.type === "table");
};

describe("tracked table cell grid provenance", () => {
  test("accepting cell removals is a save/reopen fixed point", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildSource());

    expect(reviewer.acceptAll()).toBe(3);
    const once = await reviewer.toBuffer();
    const tables = await tablesFrom(once);

    expect(tables.at(0)?.columnWidths).toEqual([4680]);
    expect(tables.at(0)?.rows.at(0)?.cells).toHaveLength(1);
    expect(tables.at(0)?.formatting?.gridSourceXml).toBe(
      '<w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>',
    );
    expect(tables.at(1)?.columnWidths).toEqual([3000, 7000]);
    expect(tables.at(1)?.formatting?.gridSourceXml).toBe(UNCHANGED_GRID_SOURCE);
    expect(gridsFrom(await documentXml(once))).toEqual([
      '<w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>',
      UNCHANGED_GRID_SOURCE,
    ]);

    const reopened = await FolioDocxReviewer.fromBuffer(once);
    const twice = await reopened.toBuffer();
    expect(await documentXml(twice)).toBe(await documentXml(once));
  });

  test("rejecting cell removals preserves both table grids verbatim", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildSource());

    expect(reviewer.rejectAll()).toBe(3);
    const once = await reviewer.toBuffer();
    const tables = await tablesFrom(once);

    expect(tables.at(0)?.columnWidths).toEqual(GRID_WIDTHS);
    expect(tables.at(0)?.rows.at(0)?.cells).toHaveLength(4);
    expect(tables.at(0)?.formatting?.gridSourceXml).toBe(GRID_SOURCE);
    expect(tables.at(1)?.columnWidths).toEqual([3000, 7000]);
    expect(tables.at(1)?.formatting?.gridSourceXml).toBe(UNCHANGED_GRID_SOURCE);
    expect(gridsFrom(await documentXml(once))).toEqual([GRID_SOURCE, UNCHANGED_GRID_SOURCE]);

    const reopened = await FolioDocxReviewer.fromBuffer(once);
    const twice = await reopened.toBuffer();
    expect(await documentXml(twice)).toBe(await documentXml(once));
  });
});
