import { describe, expect, setDefaultTimeout, test } from "bun:test";

import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));

import type { Document, Paragraph, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { tableGrid } from "../tableGrid";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS, type DocumentOp } from "../types";

const exact = (actual: unknown, expected: unknown): void => {
  expect(actual === null).toBe(expected === null);
  expect(typeof actual).toBe(typeof expected);
  if (
    actual === null ||
    expected === null ||
    typeof actual !== "object" ||
    typeof expected !== "object"
  ) {
    expect(actual).toBe(expected);
    return;
  }
  expect(Array.isArray(actual)).toBe(Array.isArray(expected));
  const keys = Object.keys(expected).sort();
  expect(Object.keys(actual).sort()).toEqual(keys);
  for (const key of keys) exact(Reflect.get(actual, key), Reflect.get(expected, key));
};

const paragraph = (paraId: string, text = "cell"): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const tableOf = (): Table => ({
  type: "table",
  formatting: { layout: "fixed", styleId: "base" },
  columnWidths: [500, 600, 700],
  rows: [
    {
      type: "tableRow",
      formatting: { header: true },
      cells: [0, 1, 2].map((column) => ({
        type: "tableCell",
        content: [paragraph(`0000001${column}`)],
      })),
    },
    {
      type: "tableRow",
      cells: [0, 1, 2].map((column) => ({
        type: "tableCell",
        content: [paragraph(`0000002${column}`)],
      })),
    },
  ],
});

const documentOf = (table: Table): Document => ({
  package: {
    document: { content: [paragraph("00000001", "before"), table, paragraph("00000002", "after")] },
  },
});
const tableIn = (document: Document): Table => {
  const block = document.package.document.content.at(1);
  if (block?.type !== "table") throw new Error("The table remains at the fixture position.");
  return block;
};
const tableTarget = { story: OP_STORIES.MAIN, blockId: "00000010" } as const;
const revision = (id: number) => ({ id, author: "Reviewer", date: "2026-10-02T10:00:00Z" });
const propertyEdit = (
  kind: "table" | "row" | "cell",
  stamp?: ReturnType<typeof revision>,
): DocumentOp => {
  switch (kind) {
    case "table":
      return {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
        patch: { styleId: "tracked" },
        ...(stamp === undefined ? {} : { revision: stamp }),
      };
    case "row":
      return {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
        patch: { cantSplit: true },
        ...(stamp === undefined ? {} : { revision: stamp }),
      };
    case "cell":
      return {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
        patch: { noWrap: true },
        ...(stamp === undefined ? {} : { revision: stamp }),
      };
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  exact(undone.value.document, document);
  return result.value;
};

const resolved = (
  document: Document,
  revisionIds: readonly number[],
  decision: "accept" | "reject",
) =>
  applied(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });

const refused = (document: Document, op: DocumentOp, reason: string) => {
  const before = structuredClone(document);
  const result = applyDocumentOp(document, op);
  if (result.isOk()) throw new Error(`Expected ${reason} refusal for ${op.type}.`);
  expect(result.error.reason).toBe(reason);
  exact(document, before);
};

describe("tracked table edits", () => {
  test.each(["table", "row", "cell"] as const)(
    "%s property revisions accept to the direct edit and reject to baseline",
    (kind) => {
      const baseline = documentOf(tableOf());
      const id = 300;
      const operation = propertyEdit(kind);
      const direct = applied(baseline, operation);
      const directlyEdited = tableIn(direct.document);
      if (kind === "table") expect(directlyEdited.formatting?.styleId).toBe("tracked");
      if (kind === "row") expect(directlyEdited.rows.at(0)?.formatting?.cantSplit).toBe(true);
      if (kind === "cell")
        expect(directlyEdited.rows.at(0)?.cells.at(0)?.formatting?.noWrap).toBe(true);
      const trackedOperation = propertyEdit(kind, revision(id));
      const tracked = applied(baseline, trackedOperation);
      const pending = tableIn(tracked.document);
      if (kind === "table")
        expect(pending.propertyChanges?.at(0)?.type).toBe("tablePropertyChange");
      if (kind === "row")
        expect(pending.rows.at(0)?.propertyChanges?.at(0)?.type).toBe("tableRowPropertyChange");
      if (kind === "cell")
        expect(pending.rows.at(0)?.cells.at(0)?.propertyChanges?.at(0)?.type).toBe(
          "tableCellPropertyChange",
        );
      const rejected = resolved(tracked.document, [id], REVISION_DECISIONS.REJECT);
      exact(rejected.document, baseline);
      const accepted = resolved(tracked.document, [id], REVISION_DECISIONS.ACCEPT);
      exact(accepted.document, direct.document);
    },
  );

  test("grid revisions restore or keep the exact column widths", () => {
    const baseline = documentOf(tableOf());
    const direct = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
      columnWidths: [610, 720, 830],
    });
    expect(tableIn(direct.document).columnWidths).toEqual([610, 720, 830]);
    const tracked = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
      columnWidths: [610, 720, 830],
      revision: revision(400),
      newIds: { revision: [401] },
    });
    expect(tableIn(tracked.document).formatting?.gridChange).toEqual({
      id: 400,
      columnWidths: [500, 600, 700],
    });
    exact(resolved(tracked.document, [400, 401], REVISION_DECISIONS.REJECT).document, baseline);
    exact(
      resolved(tracked.document, [400, 401], REVISION_DECISIONS.ACCEPT).document,
      direct.document,
    );
  });

  test.each(["absent", "empty", "sourceXml"] as const)(
    "grid resolution preserves %s table formatting exactly",
    (formattingKind) => {
      const table = tableOf();
      if (formattingKind === "absent") delete table.formatting;
      if (formattingKind === "empty") table.formatting = {};
      if (formattingKind === "sourceXml") table.formatting = { sourceXml: "<w:tblPr/>" };
      const baseline = documentOf(table);
      const operation = {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
        columnWidths: [510, 620, 730],
      } satisfies DocumentOp;
      const direct = applied(baseline, operation);
      const tracked = applied(baseline, {
        ...operation,
        revision: revision(410),
        newIds: { revision: [411] },
      } satisfies DocumentOp);
      const rejected = resolved(tracked.document, [410, 411], REVISION_DECISIONS.REJECT);
      exact(rejected.document, baseline);
      const accepted = resolved(tracked.document, [410, 411], REVISION_DECISIONS.ACCEPT);
      exact(accepted.document, direct.document);
    },
  );

  test("direct grid edits invert own undefined formatting and tracked edits refuse it", () => {
    const table = tableOf();
    table.formatting = undefined;
    const baseline = documentOf(table);
    applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
      columnWidths: [510, 620, 730],
    });
    refused(
      baseline,
      {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
        columnWidths: [510, 620, 730],
        revision: revision(420),
        newIds: { revision: [421] },
      },
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
  });

  test("tracked insertion resolves cells before the grid snapshot", () => {
    const baseline = documentOf(tableOf());
    const direct = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
      column: 1,
      width: 550,
      newBlockIds: ["00000301", "00000302"],
    });
    expect(tableIn(direct.document).columnWidths).toEqual([500, 550, 600, 700]);
    expect(tableIn(direct.document).rows.at(0)?.cells).toHaveLength(4);
    const tracked = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
      column: 1,
      width: 550,
      newBlockIds: ["00000301", "00000302"],
      revision: revision(500),
      newIds: { revision: [501, 502, 503] },
    });
    exact(
      resolved(tracked.document, [500, 501, 502, 503], REVISION_DECISIONS.REJECT).document,
      baseline,
    );
    exact(
      resolved(tracked.document, [500, 501, 502, 503], REVISION_DECISIONS.ACCEPT).document,
      direct.document,
    );
  });

  test.each(["absent", "empty"] as const)(
    "tracked column insertion preserves %s table formatting",
    (formattingKind) => {
      const table = tableOf();
      if (formattingKind === "absent") delete table.formatting;
      if (formattingKind === "empty") table.formatting = {};
      const baseline = documentOf(table);
      const op = {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        column: 1,
        width: 550,
        newBlockIds: ["00000311", "00000312"],
      } satisfies DocumentOp;
      const direct = applied(baseline, op);
      const tracked = applied(baseline, {
        ...op,
        revision: revision(430),
        newIds: { revision: [431, 432, 433] },
      } satisfies DocumentOp);
      exact(
        resolved(tracked.document, [430, 431, 432, 433], REVISION_DECISIONS.REJECT).document,
        baseline,
      );
      exact(
        resolved(tracked.document, [430, 431, 432, 433], REVISION_DECISIONS.ACCEPT).document,
        direct.document,
      );
    },
  );

  test("tracked deletion accepts to the direct grid change and rejects to baseline", () => {
    const baseline = documentOf(tableOf());
    const direct = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      column: 1,
    });
    expect(tableIn(direct.document).columnWidths).toEqual([500, 700]);
    expect(tableIn(direct.document).rows.at(0)?.cells).toHaveLength(2);
    const tracked = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      column: 1,
      revision: revision(600),
      newIds: { revision: [601, 602, 603] },
    });
    exact(
      resolved(tracked.document, [600, 601, 602, 603], REVISION_DECISIONS.REJECT).document,
      baseline,
    );
    exact(
      resolved(tracked.document, [600, 601, 602, 603], REVISION_DECISIONS.ACCEPT).document,
      direct.document,
    );
  });

  test("tracked horizontal split rejects through the owner cell property snapshot", () => {
    const baseTable = tableOf();
    const firstRow = baseTable.rows.at(0);
    if (firstRow === undefined) throw new Error("The fixture has a first row.");
    const first = firstRow.cells.at(0);
    const last = firstRow.cells.at(2);
    if (first === undefined || last === undefined) throw new Error("The fixture cells exist.");
    baseTable.rows[0] = {
      ...firstRow,
      cells: [{ ...first, formatting: { gridSpan: 2 } }, last],
    };
    const baseline = documentOf(baseTable);
    const direct = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
      newBlockIds: ["00000303"],
    });
    expect(tableIn(direct.document).rows.at(0)?.cells).toHaveLength(3);
    expect(tableIn(direct.document).rows.at(0)?.cells.at(0)?.formatting?.gridSpan).toBeUndefined();
    const tracked = applied(baseline, {
      ...tableTarget,
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
      newBlockIds: ["00000303"],
      revision: revision(650),
      newIds: { revision: [651] },
    });
    exact(resolved(tracked.document, [650, 651], REVISION_DECISIONS.REJECT).document, baseline);
    exact(
      resolved(tracked.document, [650, 651], REVISION_DECISIONS.ACCEPT).document,
      direct.document,
    );
  });

  test.each(["absent", "empty"] as const)(
    "tracked pure vertical merge resolves paired and partial revisions with %s formatting",
    (formattingKind) => {
      const table = tableOf();
      for (const row of table.rows) {
        const cell = row.cells.at(0);
        if (cell === undefined) throw new Error("The merge cells exist.");
        if (formattingKind === "empty") cell.formatting = {};
        else delete cell.formatting;
      }
      const second = table.rows.at(1)?.cells.at(0);
      if (second === undefined) throw new Error("The continuation cell exists.");
      second.content = [paragraph("00000020", "")];
      const baseline = documentOf(table);
      const direct = applied(baseline, {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        top: 0,
        bottom: 2,
        left: 0,
        right: 1,
        newBlockIds: [],
      });
      const tracked = applied(baseline, {
        ...tableTarget,
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        top: 0,
        bottom: 2,
        left: 0,
        right: 1,
        newBlockIds: [],
        revision: revision(660),
        newIds: { revision: [661, 662, 663] },
      });
      const trackedTable = tableIn(tracked.document);
      const owner = trackedTable.rows.at(0)?.cells.at(0);
      const continuation = trackedTable.rows.at(1)?.cells.at(0);
      if (owner?.structuralChange?.type !== "tableCellMerge")
        throw new Error("The owner cell has a cellMerge record.");
      if (continuation?.structuralChange?.type !== "tableCellMerge")
        throw new Error("The continuation cell has a cellMerge record.");
      const ownerProperty = owner.propertyChanges?.at(0);
      const continuationProperty = continuation.propertyChanges?.at(0);
      if (ownerProperty === undefined || continuationProperty === undefined)
        throw new Error("Both vertical cells have paired revision records.");
      const ownerMergeId = owner.structuralChange.info.id;
      const ownerPropertyId = ownerProperty.info.id;
      const continuationMergeId = continuation.structuralChange.info.id;
      const continuationPropertyId = continuationProperty.info.id;

      for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
        const propertyResolved = resolved(tracked.document, [ownerPropertyId], decision);
        const partialTable = tableIn(propertyResolved.document);
        expect(partialTable.rows.at(0)?.cells.at(0)?.structuralChange?.info.id).toBe(ownerMergeId);
        expect(partialTable.rows.at(0)?.cells.at(0)?.propertyChanges).toBeUndefined();
        expect(partialTable.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBe("restart");
        expect(tableGrid(partialTable, DOCUMENT_OP_TYPES.MERGE_CELLS).isOk()).toBe(true);

        const remaining = [ownerMergeId, continuationMergeId, continuationPropertyId];
        const final = resolved(propertyResolved.document, remaining, decision);
        if (decision === REVISION_DECISIONS.ACCEPT) exact(final.document, direct.document);
        else if (formattingKind === "absent") exact(final.document, baseline);
        else {
          // OOXML has no distinction between absent and empty formatting once its tcPrChange was resolved alone.
          const rejectedTable = tableIn(final.document);
          expect(rejectedTable.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
          expect(rejectedTable.rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
          expect(tableGrid(rejectedTable, DOCUMENT_OP_TYPES.MERGE_CELLS).isOk()).toBe(true);
        }
      }
      exact(
        resolved(
          tracked.document,
          [ownerMergeId, ownerPropertyId, continuationMergeId, continuationPropertyId],
          REVISION_DECISIONS.REJECT,
        ).document,
        baseline,
      );
      exact(
        resolved(
          tracked.document,
          [ownerMergeId, ownerPropertyId, continuationMergeId, continuationPropertyId],
          REVISION_DECISIONS.ACCEPT,
        ).document,
        direct.document,
      );
    },
  );

  test("vertical cellMerge marks resolve their current and original merge states", () => {
    const table = tableOf();
    const first = table.rows.at(0)?.cells.at(0);
    const second = table.rows.at(1)?.cells.at(0);
    if (first === undefined || second === undefined) throw new Error("The merge cells exist.");
    first.formatting = { vMerge: "restart" };
    second.formatting = { vMerge: "continue" };
    second.structuralChange = {
      type: "tableCellMerge",
      info: { id: 700, author: "Reviewer" },
      verticalMerge: "continue",
      verticalMergeOriginal: "rest",
    };
    const baseline = documentOf(table);
    const accepted = resolved(baseline, [700], REVISION_DECISIONS.ACCEPT);
    expect(tableIn(accepted.document).rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBe("continue");
    expect(tableIn(accepted.document).rows.at(1)?.cells.at(0)?.structuralChange).toBeUndefined();
    const rejected = resolved(baseline, [700], REVISION_DECISIONS.REJECT);
    expect(tableIn(rejected.document).rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBe("restart");
    expect(tableIn(rejected.document).rows.at(1)?.cells.at(0)?.structuralChange).toBeUndefined();
  });

  test("stale table inverses and incomplete cell-column revisions refuse atomically", () => {
    const baseline = documentOf(tableOf());
    const originalTable = tableIn(baseline);
    refused(
      baseline,
      {
        type: DOCUMENT_OP_TYPES.SET_TABLE,
        ...tableTarget,
        expected: { ...originalTable, columnWidths: [1, 2, 3] },
        table: originalTable,
      },
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
    );

    const malformed = structuredClone(baseline);
    const markedTable = tableIn(malformed);
    const cell = markedTable.rows.at(0)?.cells.at(1);
    if (cell === undefined) throw new Error("The target cell exists.");
    cell.structuralChange = { type: "tableCellDeletion", info: { id: 800, author: "Reviewer" } };
    refused(
      malformed,
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [800],
        decision: REVISION_DECISIONS.ACCEPT,
      },
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
  });
});

test("imported grid and table property histories resolve with the selected baseline and exact inverse", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 10, max: 10000 }),
      fc.integer({ min: 10, max: 10000 }),
      fc.integer({ min: 10, max: 10000 }),
      fc.boolean(),
      fc.boolean(),
      (originalWidth, intermediateWidth, currentWidth, reject, selectBoth) => {
        const table = tableOf();
        table.formatting = {
          width: currentWidth,
          gridChange: { id: 503, columnWidths: [510, 610, 710] },
        };
        table.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: revision(501),
            previousFormatting: { width: originalWidth },
          },
          {
            type: "tablePropertyChange",
            info: revision(502),
            previousFormatting: { width: intermediateWidth },
          },
        ];
        const before = documentOf(table);
        const result = resolved(
          before,
          selectBoth ? [501, 502, 503] : [502, 503],
          reject ? REVISION_DECISIONS.REJECT : REVISION_DECISIONS.ACCEPT,
        );
        const after = tableIn(result.document);
        let expectedWidth = currentWidth;
        if (reject) expectedWidth = selectBoth ? originalWidth : intermediateWidth;
        expect(after.formatting?.width).toBe(expectedWidth);
        expect(after.formatting?.gridChange).toBeUndefined();
        expect(after.columnWidths).toEqual(reject ? [510, 610, 710] : [500, 600, 700]);
        expect(after.propertyChanges?.map(({ info }) => info.id) ?? []).toEqual(
          selectBoth ? [] : [501],
        );
        exact(applyDocumentOps(result.document, result.inverse).unwrap().document, before);
      },
    ),
  );
});
