import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import { getTrackedChangeGroupsForIds } from "../ai-edits/read";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, Table, TableCell } from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const paragraph = (paraId: string, text?: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === undefined
      ? []
      : [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (...content: TableCell["content"]): TableCell => ({
  type: "tableCell",
  content,
});
const nested: Table = {
  type: "table",
  columnWidths: [900],
  rows: [
    { type: "tableRow", cells: [cell(paragraph("00000003", "Nested 1"))] },
    { type: "tableRow", cells: [cell(paragraph("00000004", "Nested 2"))] },
  ],
};

const nestedWithProperties = {
  ...nested,
  formatting: { justification: "center" },
  propertyChanges: [
    {
      type: "tablePropertyChange",
      info: { id: 41, author: "Earlier" },
      previousFormatting: { justification: "right" },
    },
  ],
  rows: [
    {
      type: "tableRow",
      formatting: { cantSplit: true },
      propertyChanges: [
        {
          type: "tableRowPropertyChange",
          info: { id: 42, author: "Earlier" },
          previousFormatting: { cantSplit: false },
        },
      ],
      cells: [
        {
          ...cell(paragraph("00000003", "Nested property change")),
          formatting: { width: { type: "dxa", value: 900 } },
          propertyChanges: [
            {
              type: "tableCellPropertyChange",
              info: { id: 43, author: "Earlier" },
              previousFormatting: { width: { type: "dxa", value: 700 } },
            },
          ],
        },
      ],
    },
  ],
} satisfies Table;

const nestedWithCellRecords = {
  type: "table",
  columnWidths: [900, 1100],
  rows: [
    {
      type: "tableRow",
      cells: [
        {
          ...cell(paragraph("00000003", "Earlier cell deletion")),
          structuralChange: {
            type: "tableCellDeletion",
            info: { id: 44, author: "Earlier" },
          },
        },
        {
          ...cell(paragraph("00000004", "Retained cell")),
          formatting: { width: { type: "dxa", value: 1100 } },
          propertyChanges: [
            {
              type: "tableCellPropertyChange",
              info: { id: 45, author: "Earlier" },
              previousFormatting: { width: { type: "dxa", value: 1000 } },
              previousStructuralChange: {
                type: "tableCellInsertion",
                info: { id: 46, author: "Earlier" },
              },
            },
          ],
        },
      ],
    },
  ],
} satisfies Table;

const parseSavedTable = (xml: string): Table => {
  const element = parseXmlDocument(
    xml.replace(
      "<w:tbl",
      `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
    ),
  );
  if (element === null) throw new TypeError("Expected well-formed table XML.");
  const parsed = parseTable(element, new Map(), null, null, null, null);
  if (parsed === undefined) throw new TypeError("Expected a reopened table.");
  return parsed;
};

const tableOf = (document: Document): Table => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "table") throw new TypeError("Expected a table.");
  return block;
};

const projection = (table: Table): unknown =>
  table.rows.map((row) =>
    row.cells.map((item) => ({
      span: item.formatting?.gridSpan ?? 1,
      merge: item.formatting?.vMerge,
      content: item.content.map((block) => {
        if (block.type === "table") return projection(block);
        if (block.type === "paragraph")
          return { paraId: block.paraId, content: block.content };
        return block;
      }),
    })),
  );

test.each([
  "horizontal",
  "rectangular",
  "retained",
  "horizontal with earlier properties",
  "rectangular with earlier properties",
  "rectangular with an earlier grid",
  "rectangular with earlier cell records",
] as const)(
  "a nested table and closing paragraph survive a tracked merge (%s)",
  (shape) => {
    const rectangular = shape.startsWith("rectangular") || shape === "retained";
    const retained = shape === "retained";
    const priorProperties = shape.endsWith("with earlier properties");
    const priorGrid = shape.endsWith("with an earlier grid");
    const priorCells = shape.endsWith("with earlier cell records");
    let nestedSource: Table = priorProperties ? nestedWithProperties : nested;
    if (priorCells) nestedSource = nestedWithCellRecords;
    if (priorGrid)
      nestedSource = {
        ...nested,
        formatting: { gridChange: { id: 44, columnWidths: [700] } },
      };
    const owner = cell(paragraph("00000001", retained ? undefined : "Owner"));
    const other = cell(
      paragraph("00000002", "Before"),
      nestedSource,
      paragraph("00000005"),
    );
    const first = {
      type: "tableRow",
      cells: [
        owner,
        rectangular ? cell(paragraph("00000006")) : other,
        ...(retained ? [cell(paragraph("00000009", "Outside"))] : []),
      ],
    } satisfies Table["rows"][number];
    const table: Table = {
      type: "table",
      columnWidths: [1000, 1400, ...(retained ? [1600] : [])],
      rows: [
        first,
        ...(rectangular
          ? [
              {
                type: "tableRow" as const,
                cells: [
                  cell(paragraph("00000007")),
                  other,
                  ...(retained ? [cell(paragraph("0000000A", "Outside"))] : []),
                ],
              },
            ]
          : []),
      ],
    };
    const ending = paragraph("00000008");
    const withTable = (block: Table): Document => ({
      package: { document: { content: [block, ending] } },
    });
    const original = withTable(table);
    const operation = {
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      top: 0,
      bottom: rectangular ? 2 : 1,
      left: 0,
      right: 2,
    } as const;
    const intent = { type: "table", operation } as const;
    const ids = createEditorIntentIdAllocator()(original, intent);
    const compiled = compileEditorIntent(original, {
      intent,
      mode: {
        type: "suggesting",
        revision: {
          id: ids.revisionId,
          author: "Reviewer",
          date: "2026-11-27T00:00:00Z",
        },
        newIds: ids.newIds,
      },
    }).unwrap();
    const op = compiled.ops.at(0);
    if (op?.type !== DOCUMENT_OP_TYPES.MERGE_CELLS)
      throw new TypeError("Expected a merge operation.");
    const continuations = Number(rectangular);
    expect(op.newBlockIds).toHaveLength(
      continuations + (priorProperties ? 3 : 4),
    );
    const direct = applyDocumentOp(original, {
      ...operation,
      newBlockIds: op.newBlockIds.slice(0, continuations),
    }).unwrap().document;
    const pending = applyDocumentOps(original, compiled.ops).unwrap();
    expect(
      applyDocumentOps(pending.document, pending.inverse).unwrap().document,
    ).toStrictEqual(original);

    const marked = tableOf(pending.document);
    const geometry = marked.formatting?.gridChange?.geometryChange;
    const nestedId =
      geometry?.type === "mergeCells"
        ? geometry.nestedMoves?.at(0)?.rows.at(0)
        : undefined;
    if (nestedId === undefined) throw new Error("Expected a nested row stamp.");
    const xml = serializeTable(marked, serializeParagraph);
    const rowProperties = priorProperties ? "<w:cantSplit/>" : "";
    expect(xml).toContain(`<w:trPr>${rowProperties}<w:ins w:id="${nestedId}"`);
    expect(xml.match(/<w:tbl>/gu)).toHaveLength(3);
    const parsed = parseSavedTable(xml);
    if (priorProperties || priorGrid) {
      const fact = parsed.formatting?.gridChange?.geometryChange;
      const aliases =
        fact?.type === "mergeCells"
          ? fact.nestedMoves?.at(0)?.revisions
          : undefined;
      if (aliases === undefined)
        throw new TypeError("Expected saved property handoffs.");
      expect(aliases.map(([source]) => source)).toEqual(
        priorGrid ? [44] : [41, 42, 43],
      );
      const doc = toProseDoc(withTable(parsed), {
        tableReviewGeometry: "physical",
      });
      for (const [source, temporary] of aliases)
        expect(getTrackedChangeGroupsForIds(doc, [source]).at(0)).toEqual([
          source,
          temporary,
        ]);

      const earlierTable = aliases.at(0);
      if (earlierTable === undefined)
        throw new TypeError("Expected the table property handoff.");
      const reviewed = applyDocumentOp(withTable(parsed), {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [earlierTable[1]],
        decision: REVISION_DECISIONS.REJECT,
      }).unwrap().document;
      const existing = tableOf(reviewed);
      const nestedTables = existing.rows.flatMap(({ cells }) =>
        cells.flatMap(({ content }) =>
          content.filter((block) => block.type === "table"),
        ),
      );
      expect(nestedTables).toHaveLength(2);
      if (priorGrid) {
        expect(nestedTables.map(({ columnWidths }) => columnWidths)).toEqual([
          [700],
          [700],
        ]);
        expect(
          nestedTables.every(
            ({ formatting }) => formatting?.gridChange === undefined,
          ),
        ).toBe(true);
      } else {
        expect(
          nestedTables.map(({ formatting }) => formatting?.justification),
        ).toEqual(["right", "right"]);
        expect(
          nestedTables.every(
            ({ propertyChanges }) => propertyChanges === undefined,
          ),
        ).toBe(true);
      }
      expect(existing.formatting?.gridChange?.geometryChange).toBeDefined();
    }
    if (priorCells) {
      const fact = parsed.formatting?.gridChange?.geometryChange;
      const aliases =
        fact?.type === "mergeCells"
          ? fact.nestedMoves?.at(0)?.revisions
          : undefined;
      if (aliases?.length !== 3)
        throw new TypeError("Expected the cell revision handoffs.");
      const [deletion, formatting, previous] = aliases;
      if (
        deletion === undefined ||
        formatting === undefined ||
        previous === undefined
      )
        throw new TypeError("Expected all cell revision handoffs.");
      expect(deletion[0]).toBe(44);
      expect(formatting[0]).toBe(45);
      expect(previous[0]).toBe(46);
      expect(xml).toMatch(
        new RegExp(
          `<w:tcPrChange w:id="${formatting[1]}"[^>]*><w:tcPr>.*?<w:cellIns w:id="${previous[1]}"`,
          "u",
        ),
      );
      const doc = toProseDoc(withTable(parsed), {
        tableReviewGeometry: "physical",
      });
      expect(getTrackedChangeGroupsForIds(doc, [44]).at(0)).toEqual(deletion);
      const descriptor = {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
      } as const;
      const accepted = applyDocumentOp(withTable(parsed), {
        ...descriptor,
        revisionIds: [deletion[1]],
        decision: REVISION_DECISIONS.ACCEPT,
      }).unwrap();
      expect(
        applyDocumentOps(accepted.document, accepted.inverse).unwrap().document,
      ).toStrictEqual(withTable(parsed));
      const saved = parseSavedTable(
        serializeTable(tableOf(accepted.document), serializeParagraph),
      );
      const pendingGeometry = saved.formatting?.gridChange?.geometryChange;
      const remainingAliases =
        pendingGeometry?.type === "mergeCells"
          ? pendingGeometry.nestedMoves?.at(0)?.revisions
          : undefined;
      expect(remainingAliases).toEqual([formatting, previous]);
      const rejected = applyDocumentOp(withTable(saved), {
        ...descriptor,
        revisionIds: [formatting[1]],
        decision: REVISION_DECISIONS.REJECT,
      }).unwrap().document;
      const nestedTables = tableOf(rejected).rows.flatMap(({ cells }) =>
        cells.flatMap(({ content }) =>
          content.filter((block) => block.type === "table"),
        ),
      );
      expect(nestedTables).toHaveLength(2);
      const revived: number[] = [];
      for (const item of nestedTables) {
        expect(item.rows.at(0)?.cells).toHaveLength(1);
        const retainedCell = item.rows.at(0)?.cells.at(0);
        expect(retainedCell?.formatting?.width).toEqual({
          type: "dxa",
          value: 1000,
        });
        expect(retainedCell?.propertyChanges).toBeUndefined();
        if (retainedCell?.structuralChange?.type === "tableCellInsertion")
          revived.push(retainedCell.structuralChange.info.id);
      }
      expect(revived.sort((left, right) => left - right)).toEqual(
        [...previous].sort((left, right) => left - right),
      );
      const reopenedRevival = withTable(
        parseSavedTable(serializeTable(tableOf(rejected), serializeParagraph)),
      );
      const afterEarlierReview = applyDocumentOp(reopenedRevival, {
        ...descriptor,
        revisionIds: [previous[1]],
        decision: REVISION_DECISIONS.ACCEPT,
      }).unwrap().document;
      const withoutCell = applyDocumentOp(original, {
        ...descriptor,
        revisionIds: [44],
        decision: REVISION_DECISIONS.ACCEPT,
      }).unwrap().document;
      const revivedSource = applyDocumentOp(withoutCell, {
        ...descriptor,
        revisionIds: [45],
        decision: REVISION_DECISIONS.REJECT,
      }).unwrap().document;
      const reviewedSource = applyDocumentOp(revivedSource, {
        ...descriptor,
        revisionIds: [46],
        decision: REVISION_DECISIONS.ACCEPT,
      }).unwrap().document;
      const reviewedDirect = applyDocumentOp(reviewedSource, {
        ...operation,
        newBlockIds: op.newBlockIds.slice(0, continuations),
      }).unwrap().document;
      for (const decision of [
        REVISION_DECISIONS.ACCEPT,
        REVISION_DECISIONS.REJECT,
      ]) {
        const completed = applyDocumentOp(afterEarlierReview, {
          ...descriptor,
          revisionIds: [nestedId],
          decision,
        }).unwrap().document;
        const expected =
          decision === REVISION_DECISIONS.ACCEPT
            ? reviewedDirect
            : reviewedSource;
        expect(projection(tableOf(completed))).toEqual(
          projection(tableOf(expected)),
        );
      }
    }

    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const expected =
        decision === REVISION_DECISIONS.ACCEPT ? direct : original;
      const resolve = (document: Document) =>
        applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [nestedId],
          decision,
        }).unwrap();
      const result = resolve(pending.document);
      expect(result.document).toStrictEqual(expected);
      expect(
        applyDocumentOps(result.document, result.inverse).unwrap().document,
      ).toStrictEqual(pending.document);
      expect(
        projection(tableOf(resolve(withTable(parsed)).document)),
      ).toStrictEqual(projection(tableOf(expected)));
    }

    const mismatched = structuredClone(parsed);
    const grid = mismatched.formatting?.gridChange;
    const history = grid?.geometryChange;
    const nestedMove =
      history?.type === "mergeCells" ? history.nestedMoves?.at(0) : undefined;
    if (
      grid === undefined ||
      history?.type !== "mergeCells" ||
      nestedMove === undefined
    )
      throw new Error("Expected a saved nested move.");
    const foreignId = 60_000;
    grid.geometryChange = {
      ...history,
      nestedMoves: [{ ...nestedMove, rows: [foreignId, ...nestedMove.rows] }],
    };
    const foreign: Table = {
      type: "table",
      rows: [{ type: "tableRow", cells: [cell(paragraph("000000BB"))] }],
      formatting: { gridChange: { id: foreignId, columnWidths: [500] } },
      columnWidths: [600],
    };
    const withPeer = withTable(mismatched);
    withPeer.package.document.content.splice(1, 0, foreign);
    const review = {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      decision: REVISION_DECISIONS.ACCEPT,
    } as const;
    const isolated = applyDocumentOp(withPeer, {
      ...review,
      revisionIds: [foreignId],
    }).unwrap().document;
    expect(tableOf(isolated)).toStrictEqual(mismatched);
    const reviewedForeign = isolated.package.document.content.at(1);
    if (reviewedForeign?.type !== "table")
      throw new Error("Expected a peer table.");
    expect(reviewedForeign.formatting?.gridChange).toBeUndefined();
    const stale = applyDocumentOp(withPeer, {
      ...review,
      revisionIds: [grid.id],
    });
    expect(stale.isErr() && stale.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
    );

    const [firstPair, secondPair] = nestedMove.paragraphs;
    if (firstPair === undefined || secondPair === undefined)
      throw new Error("Expected two nested paragraph pairs.");
    grid.geometryChange = {
      ...history,
      nestedMoves: [
        {
          ...nestedMove,
          paragraphs: [
            [firstPair[0], secondPair[1]],
            [secondPair[0], firstPair[1]],
          ],
        },
      ],
    };
    const crossed = applyDocumentOp(withPeer, {
      ...review,
      revisionIds: [grid.id],
    });
    expect(crossed.isErr() && crossed.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
    );
  },
);

test.each([
  "some rows",
  "every row",
  "all deleted rows",
  "omitted row",
  "fully omitted rows",
] as const)(
  "earlier nested row revisions remain independent when a merge moves %s",
  (coverage) => {
    const sourceRows = nested.rows.map((row, index) => {
      const omitted =
        coverage === "fully omitted rows" ||
        (coverage === "omitted row" && index === 1);
      const sourceRow = omitted
        ? { ...row, cells: [], formatting: { gridBefore: 1 } }
        : row;
      if (index === 1 && (omitted || coverage === "some rows"))
        return sourceRow;
      const cells =
        coverage === "every row" && index === 0
          ? [
              cell(
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: [cell(paragraph("00000007", "Deeper"))],
                    },
                  ],
                },
                paragraph("00000003", "Nested 1"),
              ),
            ]
          : sourceRow.cells;
      return {
        ...sourceRow,
        cells,
        structuralChange: {
          type:
            index === 0 || coverage === "all deleted rows"
              ? "tableRowDeletion"
              : "tableRowInsertion",
          info: { id: 51 + index, author: "Earlier" },
        },
      };
    }) satisfies Table["rows"];
    const start = paragraph("00000001", "Owner");
    const end = paragraph("00000005");
    const originalTable = {
      type: "table",
      columnWidths: [1000, 1400],
      rows: [
        {
          type: "tableRow",
          cells: [cell(start), cell({ ...nested, rows: sourceRows }, end)],
        },
      ],
    } satisfies Table;
    const withTable = (table: Table): Document => ({
      package: { document: { content: [table, paragraph("00000006")] } },
    });
    const original = withTable(originalTable);
    const operation = {
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      top: 0,
      bottom: 1,
      left: 0,
      right: 2,
    } as const;
    const intent = { type: "table", operation } as const;
    const ids = createEditorIntentIdAllocator()(original, intent);
    const compiled = compileEditorIntent(original, {
      intent,
      mode: {
        type: "suggesting",
        revision: { id: ids.revisionId, author: "Reviewer" },
        newIds: ids.newIds,
      },
    }).unwrap();
    const pending = applyDocumentOps(original, compiled.ops).unwrap();
    expect(
      applyDocumentOps(pending.document, pending.inverse).unwrap().document,
    ).toStrictEqual(original);
    const xml = serializeTable(tableOf(pending.document), serializeParagraph);
    const parsed = parseSavedTable(xml);
    const grid = parsed.formatting?.gridChange;
    const move =
      grid?.geometryChange?.type === "mergeCells"
        ? grid.geometryChange.nestedMoves?.at(0)
        : undefined;
    const aliases = move?.revisions;
    if (grid === undefined || aliases === undefined)
      throw new TypeError("Expected saved row handoffs.");
    expect(aliases.map(([source]) => source)).toEqual(
      coverage === "every row" || coverage === "all deleted rows"
        ? [51, 52]
        : [51],
    );
    const reopened = withTable(parsed);
    const doc = toProseDoc(reopened, { tableReviewGeometry: "physical" });
    for (const [source, temporary] of aliases) {
      const mark =
        source === 51 || coverage === "all deleted rows" ? "del" : "ins";
      expect(xml).toMatch(
        new RegExp(
          `<w:trPr>(?:<w:gridBefore[^>]*\\/>)?<w:${mark} w:id="${temporary}"`,
          "u",
        ),
      );
      expect(getTrackedChangeGroupsForIds(doc, [source]).at(0)).toEqual([
        source,
        temporary,
      ]);
    }
    const originalIds = aliases.map(([source]) => source);
    const temporaryIds = aliases.map(([, temporary]) => temporary);
    const resolve = (
      document: Document,
      revisionIds: number[],
      decision: "accept" | "reject",
    ): Document =>
      applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision,
      }).unwrap().document;
    const reopen = (document: Document): Document =>
      withTable(
        parseSavedTable(serializeTable(tableOf(document), serializeParagraph)),
      );
    const direct = applyDocumentOp(original, {
      ...operation,
      newBlockIds: [],
    }).unwrap().document;
    expect(
      resolve(pending.document, [grid.id], REVISION_DECISIONS.ACCEPT),
    ).toStrictEqual(direct);
    expect(
      resolve(pending.document, [grid.id], REVISION_DECISIONS.REJECT),
    ).toStrictEqual(original);
    for (const rowDecision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const decidedOriginal = resolve(original, originalIds, rowDecision);
      const earlier = reopen(resolve(reopened, temporaryIds, rowDecision));
      // A move keeps the original closing paragraph even if an independent
      // row decision later removes the entire table that preceded it.
      const decidedDirect = resolve(direct, originalIds, rowDecision);
      for (const mergeDecision of [
        REVISION_DECISIONS.ACCEPT,
        REVISION_DECISIONS.REJECT,
      ]) {
        const expected =
          mergeDecision === REVISION_DECISIONS.ACCEPT
            ? decidedDirect
            : decidedOriginal;
        const first = resolve(earlier, [grid.id], mergeDecision);
        const second = resolve(
          reopen(resolve(reopened, [grid.id], mergeDecision)),
          originalIds,
          rowDecision,
        );
        for (const completed of [first, second])
          expect(projection(tableOf(completed))).toStrictEqual(
            projection(tableOf(expected)),
          );
        if (mergeDecision === rowDecision) {
          const together = resolve(
            reopened,
            [grid.id, ...originalIds],
            rowDecision,
          );
          expect(projection(tableOf(together))).toStrictEqual(
            projection(tableOf(expected)),
          );
        }
      }
    }
  },
);

test.each(["deleteColumn", "mergeCells"] as const)(
  "a moved nested table keeps its earlier grouped geometry actionable (%s)",
  (kind) => {
    const inner = {
      type: "table",
      columnWidths: [700, 900],
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              ...cell(paragraph("00000003", "Span")),
              formatting: { gridSpan: 2 },
            },
          ],
        },
        {
          type: "tableRow",
          cells: [
            cell(paragraph("00000004", "Removed")),
            cell(
              paragraph(
                "00000007",
                kind === "deleteColumn" ? "Survivor" : undefined,
              ),
            ),
          ],
        },
      ],
    } satisfies Table;
    const source = {
      type: "table",
      columnWidths: [1000, 1400],
      rows: [
        {
          type: "tableRow",
          cells: [
            cell(paragraph("00000001", "Owner")),
            cell(inner, paragraph("00000005")),
          ],
        },
      ],
    } satisfies Table;
    const withTable = (table: Table): Document => ({
      package: { document: { content: [table, paragraph("00000006")] } },
    });
    const deletion = {
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      story: OP_STORIES.MAIN,
      blockId: "00000003",
      column: 0,
    } as const;
    const merge = {
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      top: 0,
      bottom: 1,
      left: 0,
      right: 2,
    } as const;
    const innerMerge = { ...merge, blockId: "00000004", top: 1, bottom: 2 };
    const suggest = (
      document: Document,
      operation: typeof deletion | typeof merge,
    ) => {
      const intent = { type: "table", operation } as const;
      const { revisionId, newIds } = createEditorIntentIdAllocator()(
        document,
        intent,
      );
      const { ops } = compileEditorIntent(document, {
        intent,
        mode: {
          type: "suggesting",
          revision: { id: revisionId, author: "Reviewer" },
          newIds,
        },
      }).unwrap();
      return applyDocumentOps(document, ops).unwrap().document;
    };
    const reopen = (document: Document): Document =>
      withTable(
        parseSavedTable(serializeTable(tableOf(document), serializeParagraph)),
      );
    const resolve = (
      document: Document,
      revisionIds: readonly number[],
      decision: "accept" | "reject",
    ) =>
      applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [...revisionIds],
        decision,
      }).unwrap().document;
    const nestedTables = (document: Document) =>
      tableOf(document).rows.flatMap(({ cells }) =>
        cells.flatMap(({ content }) =>
          content.filter((block) => block.type === "table"),
        ),
      );
    const earlier = suggest(
      withTable(source),
      kind === "deleteColumn" ? deletion : innerMerge,
    );
    const prior = nestedTables(earlier).at(0)?.formatting?.gridChange;
    if (prior?.geometryChange?.type !== kind)
      throw new TypeError("Expected earlier grid geometry.");

    const moved = suggest(earlier, merge);
    const reopened = reopen(moved);
    const outer = tableOf(reopened).formatting?.gridChange;
    const aliases =
      outer?.geometryChange?.type === "mergeCells"
        ? outer.geometryChange.nestedMoves?.at(0)?.revisions
        : undefined;
    const priorCopy = aliases?.find(([original]) => original === prior.id);
    if (outer === undefined || aliases === undefined || priorCopy === undefined)
      throw new TypeError("Expected grid geometry handoffs.");
    const projectionDoc = toProseDoc(reopened, {
      tableReviewGeometry: "physical",
    });
    expect(
      getTrackedChangeGroupsForIds(projectionDoc, [prior.id]).at(0),
    ).toContain(priorCopy[1]);
    const members = new Map(aliases);
    expect(
      nestedTables(reopened).map(
        ({ formatting }) => formatting?.gridChange?.geometryChange?.revisionIds,
      ),
    ).toEqual(
      expect.arrayContaining([
        prior.geometryChange.revisionIds,
        prior.geometryChange.revisionIds.map((id) => members.get(id)),
      ]),
    );
    const directMerge = (document: Document) =>
      applyDocumentOp(document, {
        ...merge,
        newBlockIds: [],
      }).unwrap().document;
    expect(resolve(moved, [outer.id], REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      directMerge(earlier),
    );
    expect(resolve(moved, [outer.id], REVISION_DECISIONS.REJECT)).toStrictEqual(
      earlier,
    );
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const reviewed = reopen(resolve(reopened, [priorCopy[1]], decision));
      const expected = resolve(earlier, [prior.id], decision);
      const widths = nestedTables(expected).at(0)?.columnWidths;
      expect(nestedTables(reviewed)).toHaveLength(2);
      for (const table of nestedTables(reviewed)) {
        expect(table.formatting?.gridChange).toBeUndefined();
        expect(table.columnWidths).toEqual(widths);
      }
      const kept = resolve(reviewed, [outer.id], REVISION_DECISIONS.ACCEPT);
      const restored = resolve(reviewed, [outer.id], REVISION_DECISIONS.REJECT);
      expect(projection(tableOf(kept))).toEqual(
        projection(tableOf(directMerge(expected))),
      );
      expect(projection(tableOf(restored))).toEqual(
        projection(tableOf(expected)),
      );
      const together = resolve(reopened, [outer.id, prior.id], decision);
      expect(projection(tableOf(together))).toEqual(
        projection(
          tableOf(
            decision === REVISION_DECISIONS.ACCEPT
              ? directMerge(expected)
              : expected,
          ),
        ),
      );
    }
  },
);

test("a merge refuses before creating an unwriteable nested-table review record", () => {
  const destinationId = (index: number) =>
    (0x4000 + index).toString(16).padStart(8, "0");
  const nestedTable: Table = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          cell(
            ...Array.from({ length: 1600 }, (_, index) =>
              paragraph(destinationId(index)),
            ),
          ),
        ],
      },
    ],
  };
  const original = {
    package: {
      document: {
        content: [
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  cell(paragraph("00000001", "A")),
                  cell(nestedTable, paragraph("00000002")),
                ],
              },
            ],
          },
          paragraph("00000003"),
        ],
      },
    },
  } satisfies Document;
  const result = applyDocumentOp(original, {
    type: DOCUMENT_OP_TYPES.MERGE_CELLS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    top: 0,
    bottom: 1,
    left: 0,
    right: 2,
    newBlockIds: Array.from({ length: 1601 }, (_, index) =>
      destinationId(0x800 + index),
    ),
    revision: { id: 100, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
    newIds: { revision: [101, 102, 103, 104, 105, 106, 107] },
  });
  expect(result.isErr()).toBe(true);
  if (result.isOk()) throw new Error("Expected a size refusal.");
  expect(result.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
  );
  expect(result.error.message).toContain("exceeds the saved history limit");
});
