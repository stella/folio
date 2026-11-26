import { describe, expect, test } from "bun:test";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import type { Document, SdtProperties, Table } from "../../types/document";
import { serializeTable } from "./tableSerializer";

const serializeParagraph = (): string => "<w:p/>";

const cell = (gridSpan?: number) =>
  ({
    type: "tableCell",
    ...(gridSpan === undefined ? {} : { formatting: { gridSpan } }),
    content: [{ type: "paragraph", content: [] }],
  }) as const satisfies Table["rows"][number]["cells"][number];

// `w:tbl` is `w:tblPr, w:tblGrid, (rows)*`: both are required and both precede
// every row. A table the model carries no properties or measured widths for —
// the shape a tracked table insertion produces — used to open with its first
// `w:tr`, which the content model has no place for.
describe("serializeTable required children", () => {
  test("opens with the properties and grid even when the model carries neither", () => {
    const table: Table = {
      type: "table",
      rows: [
        { type: "tableRow", cells: [cell(), cell()] },
        { type: "tableRow", cells: [cell(), cell()] },
      ],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(
      xml.startsWith(
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>",
      ),
    ).toBe(true);
  });

  test("declares a grid column for every column the widest row spans", () => {
    const table: Table = {
      type: "table",
      rows: [
        { type: "tableRow", cells: [cell(3)] },
        { type: "tableRow", cells: [cell(), cell(), cell()] },
      ],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml.match(/<w:gridCol\b/gu)).toHaveLength(3);
  });

  test("keeps authored column widths when the model has them", () => {
    const table: Table = {
      type: "table",
      columnWidths: [2000, 3000],
      rows: [{ type: "tableRow", cells: [cell(), cell()] }],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml).toContain(
      '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>',
    );
  });
});

// A grid captured from the source is written back verbatim only while it
// still declares the grid the rows cover: an edit that adds a column without
// measuring it leaves the widths unchanged (none) but not the grid.
describe("serializeTable grid source", () => {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const twoColumns = `<w:tblGrid ${W}><w:gridCol/><w:gridCol/></w:tblGrid>`;

  test("keeps the captured grid while it still describes the table", () => {
    const table: Table = {
      type: "table",
      formatting: { gridSourceXml: twoColumns },
      rows: [{ type: "tableRow", cells: [cell(), cell()] }],
    };

    expect(serializeTable(table, serializeParagraph)).toContain(twoColumns);
  });

  test("rebuilds it once the rows cover another column and a grid change", () => {
    const table: Table = {
      type: "table",
      formatting: {
        gridSourceXml: twoColumns,
        gridChange: { id: 4, columnWidths: [undefined, undefined] },
      },
      rows: [{ type: "tableRow", cells: [cell(), cell(), cell()] }],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml).toContain(
      '<w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/><w:tblGridChange w:id="4"><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid></w:tblGridChange></w:tblGrid>',
    );
  });

  test("declares the columns a row skips before and after its cells", () => {
    const table: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          formatting: { gridBefore: 1, gridAfter: 1 },
          cells: [cell()],
        },
        {
          type: "tableRow",
          formatting: { gridBefore: 2 },
          cells: [cell()],
        },
      ],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml.match(/<w:gridCol\b/gu)).toHaveLength(3);
  });
});

describe("serializeTable shared content controls after edits", () => {
  const control = { id: 4242, sdtType: "richText" } satisfies SdtProperties;
  const wrapped = { contentControls: [control] };
  const captureXml = '<w:proofErr w:type="spellStart"/>';
  const first = "00000001";
  const second = "00000002";
  const outside = "00000003";
  const inserted = "00000004";
  const anchoredCell = (paraId: string) => ({
    type: "tableCell" as const,
    content: [{ type: "paragraph" as const, paraId, content: [] }],
  });
  const row = (paraId: string) => ({
    type: "tableRow" as const,
    cells: [anchoredCell(paraId)],
  });
  const xmlOf = (document: Document): string => {
    const table = document.package.document.content.at(0);
    if (table?.type !== "table") throw new Error("Expected fixture table.");
    return serializeTable(
      table,
      ({ paraId }) => `<w:p w14:paraId="${paraId ?? ""}"/>`,
    );
  };
  const applied = (document: Document, op: DocumentOp) => {
    const result = applyDocumentOp(document, op);
    if (result.isErr()) throw result.error;
    return result.value;
  };

  for (const scope of ["row", "cell"] as const) {
    test.each(["inside", "after"] as const)(
      `${scope} insertion %s a shared control keeps the capture in one wrapper`,
      (placement) => {
        const at = placement === "inside" ? 1 : 2;
        const preserved = {
          children: [{ ...wrapped, index: at, xml: captureXml }],
        };
        const table: Table =
          scope === "row"
            ? {
                type: "table",
                preserved,
                rows: [
                  { ...row(first), ...wrapped },
                  { ...row(second), ...wrapped },
                  row(outside),
                ],
              }
            : {
                type: "table",
                columnWidths: [500, 600, 700],
                rows: [
                  {
                    type: "tableRow",
                    preserved,
                    cells: [
                      { ...anchoredCell(first), ...wrapped },
                      { ...anchoredCell(second), ...wrapped },
                      anchoredCell(outside),
                    ],
                  },
                ],
              };
        const baseline = {
          package: { document: { content: [table] } },
        } satisfies Document;
        const target = { story: OP_STORIES.MAIN, blockId: first };
        const operation =
          scope === "row"
            ? {
                ...target,
                type: DOCUMENT_OP_TYPES.INSERT_ROW,
                at,
                row: row(inserted),
              }
            : {
                ...target,
                type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
                column: at,
                width: 800,
                newBlockIds: [inserted],
              };
        const direct = applied(baseline, operation);
        const tracked = applied(baseline, {
          ...operation,
          revision: {
            id: 100,
            author: "Reviewer",
            date: "2026-11-26T00:00:00Z",
          },
          newIds: { revision: [101, 102, 103] },
        });

        for (const edit of [direct, tracked]) {
          const xml = xmlOf(edit.document);
          const wrapper = /<w:sdtContent>(.*?)<\/w:sdtContent>/u.exec(xml)?.[1];
          expect(xml.match(/<w:sdt\b/gu)).toHaveLength(1);
          if (wrapper === undefined) throw new Error("Expected SDT content.");
          expect(wrapper).toContain(captureXml);
          expect(wrapper).toContain(`w14:paraId="${first}"`);
          expect(wrapper).toContain(`w14:paraId="${second}"`);
          expect(wrapper).not.toContain(`w14:paraId="${outside}"`);
          expect(xml).toContain(`w14:paraId="${inserted}"`);
          expect(wrapper.includes(`w14:paraId="${inserted}"`)).toBe(
            placement === "inside",
          );

          const undone = applyDocumentOps(edit.document, edit.inverse);
          if (undone.isErr()) throw undone.error;
          expect(xmlOf(undone.value.document)).toBe(xmlOf(baseline));
        }

        expect(xmlOf(tracked.document)).toContain(
          scope === "row" ? '<w:ins w:id="100"' : '<w:cellIns w:id="100"',
        );
        const revisionIds = tracked.revisions;
        if (revisionIds === undefined)
          throw new Error("Expected revision ids.");
        for (const decision of [
          REVISION_DECISIONS.ACCEPT,
          REVISION_DECISIONS.REJECT,
        ]) {
          const resolved = applied(tracked.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds,
            decision,
          });
          expect(xmlOf(resolved.document)).toBe(
            xmlOf(
              decision === REVISION_DECISIONS.ACCEPT
                ? direct.document
                : baseline,
            ),
          );
        }
      },
    );
  }
});
