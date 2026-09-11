import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";

import { folioStoryTables } from "../../ai-edits/snapshot";
import { resolveAllChangesInHeadlessState } from "../../prosemirror/commands/comments";
import { schema } from "../../prosemirror/schema";
import {
  DEFAULT_TABLE_GEOMETRY_PREFLIGHT_LIMITS,
  executeTableGeometryProgram,
  preflightTableGeometry,
  projectTableGeometry,
  type TableCellCoordinate,
  type TableGeometryPairing,
} from "./table-geometry-program";

const REVISION = {
  author: "Reviewer",
  date: "2026-09-11T12:00:00.000Z",
  idSeed: 40,
} as const;

const paragraph = (text: string): PMNode =>
  schema.node("paragraph", null, text.length === 0 ? [] : [schema.text(text)]);

const cell = (text: string, attrs: Record<string, unknown> = {}): PMNode =>
  schema.node("tableCell", attrs, [paragraph(text)]);

const row = (
  cells: readonly PMNode[],
  attrs: Record<string, unknown> = {},
): PMNode => schema.node("tableRow", attrs, cells);

const table = (
  rows: readonly PMNode[],
  attrs: Record<string, unknown> = {},
): PMNode => schema.node("table", attrs, rows);

const documentWith = (...blocks: readonly PMNode[]): PMNode => schema.node("doc", null, blocks);

const coordinate = (
  tableIndex: number,
  rowIndex = 0,
  cellIndex = 0,
): TableCellCoordinate => ({ tableIndex, rowIndex, cellIndex });

const pairing = (
  base: TableCellCoordinate = coordinate(0),
  target: TableCellCoordinate = coordinate(0),
): TableGeometryPairing => ({ base, target });

const targetTablesOf = (doc: PMNode): ReadonlyMap<number, PMNode> =>
  new Map(folioStoryTables(doc).map(({ index, node }) => [index, node] as const));

const preflight = (
  base: PMNode,
  target: PMNode,
  pairings: readonly TableGeometryPairing[] = [pairing()],
) =>
  preflightTableGeometry({
    baseTables: folioStoryTables(base),
    targetTables: targetTablesOf(target),
    pairings,
  });

const changedDocuments = () => {
  const base = documentWith(
    table(
      [
        row(
          [
            cell("Terms", {
              width: 2400,
              widthType: "dxa",
              backgroundColor: "FFF2CC",
            }),
          ],
          { height: 320, heightRule: "exact" },
        ),
      ],
      { width: 4800, widthType: "dxa", justification: "left", columnWidths: [4800] },
    ),
  );
  const target = documentWith(
    table(
      [
        row(
          [
            cell("Terms", {
              width: 3000,
              widthType: "dxa",
              backgroundColor: "C6E0B4",
              verticalAlign: "center",
            }),
          ],
          { height: 480, heightRule: "atLeast", isHeader: true },
        ),
      ],
      { width: 6000, widthType: "dxa", justification: "center", columnWidths: [4800] },
    ),
  );
  return { base, target };
};

const applyProgram = (
  base: PMNode,
  program: Extract<ReturnType<typeof preflight>, { status: "ready" }>["program"],
) => {
  const state = EditorState.create({ schema, doc: base });
  const transaction = state.tr;
  const execution = executeTableGeometryProgram({ tr: transaction, program, revision: REVISION });
  if (execution.status === "unsupported") {
    throw new Error(`unexpected execution refusal: ${execution.issue.reason}`);
  }
  return { state: state.apply(transaction), transaction, receipt: execution.receipt };
};

describe("atomic table geometry programs", () => {
  test("classifies exact input as an immutable unchanged program", () => {
    const base = documentWith(table([row([cell("same")])], { columnWidths: [2400] }));
    const target = documentWith(table([row([cell("same")])], { columnWidths: [2400] }));

    const result = preflight(base, target);

    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    expect(result.program.type).toBe("unchanged");
    expect(result.program.instructions).toEqual([]);
    expect(Object.isFrozen(result.program)).toBe(true);
    expect(Object.isFrozen(result.program.instructions)).toBe(true);

    const { transaction, receipt } = applyProgram(base, result.program);
    expect(transaction.steps).toHaveLength(0);
    expect(receipt).toEqual({
      type: "table-geometry-execution",
      startingRevisionId: 40,
      nextRevisionId: 40,
      revisions: [],
    });
  });

  test("accepting reconstructs the target and rejecting reconstructs the base", () => {
    const { base, target } = changedDocuments();
    const result = preflight(base, target);

    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    expect(result.program.type).toBe("changes");
    const { state, receipt } = applyProgram(base, result.program);
    expect(receipt.revisions.map(({ scope, revisionId }) => ({ scope, revisionId }))).toEqual([
      { scope: "table", revisionId: 40 },
      { scope: "row", revisionId: 41 },
      { scope: "cell", revisionId: 42 },
    ]);
    expect(receipt.nextRevisionId).toBe(43);

    const accepted = resolveAllChangesInHeadlessState(state, "accept");
    const rejected = resolveAllChangesInHeadlessState(state, "reject");
    expect(projectTableGeometry(folioStoryTables(accepted.doc))).toEqual(
      projectTableGeometry(folioStoryTables(target)),
    );
    expect(projectTableGeometry(folioStoryTables(rejected.doc))).toEqual(
      projectTableGeometry(folioStoryTables(base)),
    );
  });

  test("detects formatting held only by the complete property payload", () => {
    const formatting = (layout: "fixed" | "autofit") => ({
      width: { value: 4800, type: "dxa" as const },
      layout,
    });
    const base = documentWith(
      table([row([cell("same")])], {
        width: 4800,
        widthType: "dxa",
        columnWidths: [4800],
        _originalFormatting: formatting("fixed"),
      }),
    );
    const target = documentWith(
      table([row([cell("same")])], {
        width: 4800,
        widthType: "dxa",
        columnWidths: [4800],
        _originalFormatting: formatting("autofit"),
      }),
    );

    const result = preflight(base, target);

    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    expect(result.program.instructions.map(({ carrier }) => carrier.scope)).toEqual(["table"]);
    const { state } = applyProgram(base, result.program);
    expect(
      projectTableGeometry(
        folioStoryTables(resolveAllChangesInHeadlessState(state, "accept").doc),
      ),
    ).toEqual(projectTableGeometry(folioStoryTables(target)));
    expect(
      projectTableGeometry(
        folioStoryTables(resolveAllChangesInHeadlessState(state, "reject").doc),
      ),
    ).toEqual(projectTableGeometry(folioStoryTables(base)));
  });

  test("owns target payload instead of consulting target nodes during execution", () => {
    const { base, target } = changedDocuments();
    const expectedTarget = schema.nodeFromJSON(target.toJSON());
    const result = preflight(base, target);
    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;

    const targetCell = folioStoryTables(target).at(0)?.node.firstChild?.firstChild;
    if (!targetCell) throw new Error("expected target cell");
    targetCell.attrs["backgroundColor"] = "000000";

    const { state } = applyProgram(base, result.program);
    const accepted = resolveAllChangesInHeadlessState(state, "accept");
    expect(projectTableGeometry(folioStoryTables(accepted.doc))).toEqual(
      projectTableGeometry(folioStoryTables(expectedTarget)),
    );
  });
});

describe("table geometry refusal boundaries", () => {
  test("reports every missing target level instead of treating it as unchanged", () => {
    const base = documentWith(table([row([cell("base")])]));
    const target = documentWith(table([row([cell("target")])]));
    const cases = [
      {
        name: "table",
        targetTables: new Map<number, PMNode>(),
        pairings: [pairing()],
        reason: "missing-table",
      },
      {
        name: "row",
        targetTables: targetTablesOf(target),
        pairings: [pairing(coordinate(0), coordinate(0, 1))],
        reason: "missing-row",
      },
      {
        name: "cell",
        targetTables: targetTablesOf(target),
        pairings: [pairing(coordinate(0), coordinate(0, 0, 1))],
        reason: "missing-cell",
      },
    ] as const;

    for (const candidate of cases) {
      const result = preflightTableGeometry({
        baseTables: folioStoryTables(base),
        targetTables: candidate.targetTables,
        pairings: candidate.pairings,
      });
      expect(result.status, candidate.name).toBe("unsupported");
      if (result.status === "ready") continue;
      expect(result.issue.reason, candidate.name).toBe(candidate.reason);
    }
  });

  test("rejects duplicate and cross-paired coordinates deterministically", () => {
    const base = documentWith(
      table([row([cell("a"), cell("b")])]),
      table([row([cell("c")])]),
    );
    const target = documentWith(
      table([row([cell("a"), cell("b")])]),
      table([row([cell("c")])]),
    );
    const duplicateBase = preflight(base, target, [
      pairing(coordinate(0, 0, 0), coordinate(0, 0, 0)),
      pairing(coordinate(0, 0, 0), coordinate(0, 0, 1)),
    ]);
    const duplicateTarget = preflight(base, target, [
      pairing(coordinate(0, 0, 0), coordinate(0, 0, 0)),
      pairing(coordinate(0, 0, 1), coordinate(0, 0, 0)),
    ]);
    const crossTable = preflight(base, target, [
      pairing(coordinate(0), coordinate(0)),
      pairing(coordinate(0, 0, 1), coordinate(1)),
    ]);

    expect(duplicateBase.status).toBe("unsupported");
    expect(duplicateTarget.status).toBe("unsupported");
    expect(crossTable.status).toBe("unsupported");
    if (
      duplicateBase.status === "unsupported" &&
      duplicateTarget.status === "unsupported" &&
      crossTable.status === "unsupported"
    ) {
      expect(duplicateBase.issue.reason).toBe("duplicate-pairing");
      expect(duplicateTarget.issue.reason).toBe("duplicate-pairing");
      expect(crossTable.issue.reason).toBe("conflicting-table-pairing");
    }
  });

  test("rejects duplicate table indexes and positions before resolving pairings", () => {
    const doc = documentWith(table([row([cell("a")])]), table([row([cell("b")])]));
    const tables = folioStoryTables(doc);
    const first = tables.at(0);
    const second = tables.at(1);
    if (!first || !second) throw new Error("expected two tables");
    const targetTables = targetTablesOf(doc);
    const duplicateIndex = preflightTableGeometry({
      baseTables: [first, { ...second, index: first.index }],
      targetTables,
      pairings: [],
    });
    const duplicatePosition = preflightTableGeometry({
      baseTables: [first, { ...second, start: first.start }],
      targetTables,
      pairings: [],
    });

    expect(duplicateIndex).toMatchObject({
      status: "unsupported",
      issue: { reason: "duplicate-table-index" },
    });
    expect(duplicatePosition).toMatchObject({
      status: "unsupported",
      issue: { reason: "duplicate-table-position" },
    });
  });

  test("reports a property change whose rejected view cannot be reconstructed", () => {
    const base = documentWith(
      table([
        row([
          cell("same", {
            noWrap: true,
          }),
        ]),
      ]),
    );
    const target = documentWith(
      table([
        row([
          cell("same"),
        ]),
      ]),
    );

    const result = preflight(base, target);

    expect(projectTableGeometry(folioStoryTables(base))).not.toEqual(
      projectTableGeometry(folioStoryTables(target)),
    );
    expect(result.status).toBe("unsupported");
    if (result.status === "ready") return;
    expect(result.issue).toMatchObject({
      reason: "non-reconstructable-property-change",
      reconstruction: "rejected",
      scope: "cell",
    });
  });

  test("reports a paired table grid change that has no tracked-change representation", () => {
    const base = documentWith(table([row([cell("same")])], { columnWidths: [2400] }));
    const target = documentWith(table([row([cell("same")])], { columnWidths: [3600] }));

    const result = preflight(base, target);

    expect(result.status).toBe("unsupported");
    if (result.status === "ready") return;
    expect(result.issue).toMatchObject({
      reason: "non-reconstructable-structure-change",
      scope: "table",
      property: "column-widths",
    });
  });

  test("bounds pairings, visited nodes, changes, and captured payload", () => {
    const { base, target } = changedDocuments();
    const limits = DEFAULT_TABLE_GEOMETRY_PREFLIGHT_LIMITS;
    const cases = [
      { limit: "maxPairings", value: 0 },
      { limit: "maxVisitedNodes", value: 0 },
      { limit: "maxChanges", value: 0 },
      { limit: "maxPayloadUnits", value: 0 },
    ] as const;
    for (const candidate of cases) {
      const result = preflightTableGeometry({
        baseTables: folioStoryTables(base),
        targetTables: targetTablesOf(target),
        pairings: [pairing()],
        limits: { ...limits, [candidate.limit]: candidate.value },
      });
      expect(result.status, candidate.limit).toBe("unsupported");
      if (result.status === "ready") continue;
      expect(result.issue.reason, candidate.limit).toBe("limit-exceeded");
      if (result.issue.reason === "limit-exceeded") {
        expect(result.issue.limit, candidate.limit).toBe(candidate.limit);
      }
    }
  });

  test("validates every live carrier before adding the first transaction step", () => {
    const { base, target } = changedDocuments();
    const result = preflight(base, target);
    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;

    const staleCell = cell("Terms", {
      width: 2400,
      widthType: "dxa",
      backgroundColor: "000000",
    });
    const stale = documentWith(
      table(
        [row([staleCell], { height: 320, heightRule: "exact" })],
        { width: 4800, widthType: "dxa", justification: "left", columnWidths: [4800] },
      ),
    );
    const transaction = EditorState.create({ schema, doc: stale }).tr;

    const execution = executeTableGeometryProgram({
      tr: transaction,
      program: result.program,
      revision: REVISION,
    });

    expect(execution).toMatchObject({
      status: "unsupported",
      issue: { reason: "stale-live-node", scope: "cell" },
    });
    expect(transaction.steps).toHaveLength(0);
    expect(transaction.doc.eq(stale)).toBe(true);
  });

  test("reports a missing live carrier without mutating the caller transaction", () => {
    const base = documentWith(paragraph("before"), table([row([cell("same")])]));
    const target = documentWith(paragraph("before"), table([row([cell("changed", { width: 1 })])]));
    const result = preflight(base, target);
    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    const transaction = EditorState.create({ schema, doc: documentWith(paragraph("short")) }).tr;

    const execution = executeTableGeometryProgram({
      tr: transaction,
      program: result.program,
      revision: REVISION,
    });

    expect(execution).toMatchObject({ status: "unsupported", issue: { reason: "missing-live-node" } });
    expect(transaction.steps).toHaveLength(0);
  });

  test("validates unchanged carriers so drift cannot bypass an empty program", () => {
    const base = documentWith(table([row([cell("same", { width: 1200 })])]));
    const target = documentWith(table([row([cell("same", { width: 1200 })])]));
    const result = preflight(base, target);
    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    expect(result.program.type).toBe("unchanged");

    const stale = documentWith(table([row([cell("same", { width: 2400 })])]));
    const transaction = EditorState.create({ schema, doc: stale }).tr;
    const execution = executeTableGeometryProgram({
      tr: transaction,
      program: result.program,
      revision: REVISION,
    });

    expect(execution).toMatchObject({
      status: "unsupported",
      issue: { reason: "stale-live-node", scope: "cell" },
    });
    expect(transaction.steps).toHaveLength(0);
    expect(transaction.doc.eq(stale)).toBe(true);
  });

  test("rejects cyclic property payloads without recursing or throwing", () => {
    const cyclic: Record<string, unknown> = {};
    cyclic["self"] = cyclic;
    const malformedTable = table([row([cell("same")])]);
    malformedTable.attrs["_originalFormatting"] = cyclic;
    const base = documentWith(malformedTable);
    const target = documentWith(table([row([cell("same")])]));

    const result = preflight(base, target);

    expect(result.status).toBe("unsupported");
    if (result.status === "ready") return;
    expect(result.issue).toMatchObject({
      reason: "invalid-property-payload",
      scope: "table",
    });
  });
});

describe("table geometry ordering", () => {
  test("allocates revisions in canonical target order independent of pairing order", () => {
    const base = documentWith(
      table([row([cell("A", { width: 1000, widthType: "dxa" })])]),
      table([row([cell("B", { width: 2000, widthType: "dxa" })])]),
    );
    const target = documentWith(
      table([row([cell("B", { width: 2200, widthType: "dxa" })])]),
      table([row([cell("A", { width: 1200, widthType: "dxa" })])]),
    );
    const reversed = [
      pairing(coordinate(0), coordinate(1)),
      pairing(coordinate(1), coordinate(0)),
    ];
    const forward = [...reversed].reverse();
    const first = preflight(base, target, reversed);
    const second = preflight(base, target, forward);
    expect(first.status).toBe("ready");
    expect(second.status).toBe("ready");
    if (first.status === "unsupported" || second.status === "unsupported") return;

    const firstReceipt = applyProgram(base, first.program).receipt;
    const secondReceipt = applyProgram(base, second.program).receipt;
    expect(firstReceipt).toEqual(secondReceipt);
    expect(firstReceipt.revisions.map(({ base: source, target: destination }) => ({
      base: source.tableIndex,
      target: destination.tableIndex,
    }))).toEqual([
      { base: 1, target: 0 },
      { base: 0, target: 1 },
    ]);
  });
});
