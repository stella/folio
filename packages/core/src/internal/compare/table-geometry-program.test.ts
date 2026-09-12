import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { folioStoryTables } from "../../ai-edits/snapshot";
import { resolveAllChangesInHeadlessState } from "../../prosemirror/commands/comments";
import {
  tableCellRejectAttrPatch,
  tableRejectAttrPatch,
  tableRowRejectAttrPatch,
} from "../../prosemirror/commands/propertyChangeScope";
import { schema } from "../../prosemirror/schema";
import type {
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
} from "../../types/document";
import type {
  CompareTableCellFormattingPropertyName,
  CompareTableFormattingPropertyName,
  CompareTableRowFormattingPropertyName,
} from "../../compare/table-format-properties";
import {
  DEFAULT_TABLE_GEOMETRY_PREFLIGHT_LIMITS,
  executeTableGeometryProgram,
  preflightTableGeometry,
  preflightTableGeometryComponents,
  projectTableGeometry,
  tableGeometryProgramSemanticChangeOccurrences,
  tableGeometryProgramTableGridTransitions,
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

const row = (cells: readonly PMNode[], attrs: Record<string, unknown> = {}): PMNode =>
  schema.node("tableRow", attrs, cells);

const table = (rows: readonly PMNode[], attrs: Record<string, unknown> = {}): PMNode =>
  schema.node("table", attrs, rows);

const documentWith = (...blocks: readonly PMNode[]): PMNode => schema.node("doc", null, blocks);

const coordinate = (tableIndex: number, rowIndex = 0, cellIndex = 0): TableCellCoordinate => ({
  tableIndex,
  rowIndex,
  cellIndex,
});

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

const BASE_TABLE_PROPERTIES = {
  width: { value: 4_800, type: "dxa" },
  justification: "left",
  cellSpacing: { value: 20, type: "dxa" },
  indent: { value: 120, type: "dxa" },
  borders: { top: { style: "dashed", size: 4 } },
  cellMargins: { top: { value: 40, type: "dxa" } },
  layout: "autofit",
  styleId: "BaseTable",
  look: { lastRow: true },
  shading: { fill: { rgb: "FFFFFF" } },
  overlap: "overlap",
  floating: { horzAnchor: "margin", vertAnchor: "text", tblpX: 100 },
  bidi: false,
} as const satisfies Required<Pick<TableFormatting, CompareTableFormattingPropertyName>>;

const TARGET_TABLE_PROPERTIES = {
  width: { value: 6_000, type: "dxa" },
  justification: "center",
  cellSpacing: { value: 40, type: "dxa" },
  indent: { value: 240, type: "dxa" },
  borders: { top: { style: "single", size: 8 } },
  cellMargins: { top: { value: 80, type: "dxa" } },
  layout: "fixed",
  styleId: "TargetTable",
  look: { firstRow: true },
  shading: { fill: { rgb: "FFF2CC" } },
  overlap: "never",
  floating: { horzAnchor: "page", vertAnchor: "margin", tblpX: 200 },
  bidi: true,
} as const satisfies Required<Pick<TableFormatting, CompareTableFormattingPropertyName>>;

const BASE_ROW_PROPERTIES = {
  gridBefore: 0,
  widthBefore: { value: 100, type: "dxa" },
  gridAfter: 0,
  widthAfter: { value: 100, type: "dxa" },
  height: { value: 240, type: "dxa" },
  heightRule: "exact",
  header: false,
  cantSplit: false,
  justification: "left",
  hidden: false,
  conditionalFormat: { lastRow: true },
} as const satisfies Required<Pick<TableRowFormatting, CompareTableRowFormattingPropertyName>>;

const TARGET_ROW_PROPERTIES = {
  gridBefore: 1,
  widthBefore: { value: 200, type: "dxa" },
  gridAfter: 1,
  widthAfter: { value: 200, type: "dxa" },
  height: { value: 480, type: "dxa" },
  heightRule: "atLeast",
  header: true,
  cantSplit: true,
  justification: "center",
  hidden: true,
  conditionalFormat: { firstRow: true },
} as const satisfies Required<Pick<TableRowFormatting, CompareTableRowFormattingPropertyName>>;

const BASE_CELL_PROPERTIES = {
  width: { value: 2_400, type: "dxa" },
  borders: { top: { style: "dashed", size: 4 } },
  margins: { top: { value: 40, type: "dxa" } },
  shading: { fill: { rgb: "FFFFFF" } },
  verticalAlign: "top",
  textDirection: "lr",
  fitText: false,
  noWrap: false,
  hideMark: false,
  conditionalFormat: { lastColumn: true },
} as const satisfies Required<Pick<TableCellFormatting, CompareTableCellFormattingPropertyName>>;

const TARGET_CELL_PROPERTIES = {
  width: { value: 3_000, type: "dxa" },
  borders: { top: { style: "single", size: 8 } },
  margins: { top: { value: 80, type: "dxa" } },
  shading: { fill: { rgb: "C6E0B4" } },
  verticalAlign: "center",
  textDirection: "tbRl",
  fitText: true,
  noWrap: true,
  hideMark: true,
  conditionalFormat: { firstColumn: true },
} as const satisfies Required<Pick<TableCellFormatting, CompareTableCellFormattingPropertyName>>;

const documentWithFormatting = ({
  tableFormatting,
  rowFormatting,
  cellFormatting,
}: {
  readonly tableFormatting: TableFormatting;
  readonly rowFormatting: TableRowFormatting;
  readonly cellFormatting: TableCellFormatting;
}): PMNode =>
  documentWith(
    table(
      [
        row(
          [cell("Terms", tableCellRejectAttrPatch(cellFormatting, cellFormatting))],
          tableRowRejectAttrPatch(rowFormatting),
        ),
      ],
      {
        ...tableRejectAttrPatch(tableFormatting),
        columnWidths: [4_800],
      },
    ),
  );

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

const expectRecursivelyFrozen = (value: unknown): void => {
  if (typeof value !== "object" || value === null) return;
  expect(Object.isFrozen(value)).toBe(true);
  for (const child of Object.values(value)) expectRecursivelyFrozen(child);
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

  test("reports every modeled property through its scope-specific semantic owner", () => {
    const base = documentWithFormatting({
      tableFormatting: {
        ...BASE_TABLE_PROPERTIES,
        gridSourceXml: "<w:tblGrid data-side='base'/>",
        sourceXml: "<w:tblPr data-side='base'/>",
      },
      rowFormatting: { ...BASE_ROW_PROPERTIES, sourceXml: "<w:trPr data-side='base'/>" },
      cellFormatting: { ...BASE_CELL_PROPERTIES, sourceXml: "<w:tcPr data-side='base'/>" },
    });
    const target = documentWithFormatting({
      tableFormatting: {
        ...TARGET_TABLE_PROPERTIES,
        gridSourceXml: "<w:tblGrid data-side='target'/>",
        sourceXml: "<w:tblPr data-side='target'/>",
      },
      rowFormatting: { ...TARGET_ROW_PROPERTIES, sourceXml: "<w:trPr data-side='target'/>" },
      cellFormatting: { ...TARGET_CELL_PROPERTIES, sourceXml: "<w:tcPr data-side='target'/>" },
    });
    const result = preflight(base, target);

    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    const occurrences = tableGeometryProgramSemanticChangeOccurrences(result.program);
    expect(occurrences).toHaveLength(3);
    expect(occurrences.map(({ change }) => change.scope)).toEqual(["table", "row", "cell"]);

    const tableChange = occurrences.at(0)?.change;
    const rowChange = occurrences.at(1)?.change;
    const cellChange = occurrences.at(2)?.change;
    expect(tableChange?.scope).toBe("table");
    expect(rowChange?.scope).toBe("row");
    expect(cellChange?.scope).toBe("cell");
    if (
      tableChange?.scope !== "table" ||
      rowChange?.scope !== "row" ||
      cellChange?.scope !== "cell"
    ) {
      return;
    }

    expect(tableChange.base).toEqual({ tableIndex: 0 });
    expect(tableChange.target).toEqual({ tableIndex: 0 });
    expect(rowChange.base).toEqual({ tableIndex: 0, rowIndex: 0 });
    expect(rowChange.target).toEqual({ tableIndex: 0, rowIndex: 0 });
    expect(cellChange.base).toEqual({ tableIndex: 0, rowIndex: 0, cellIndex: 0 });
    expect(cellChange.target).toEqual({ tableIndex: 0, rowIndex: 0, cellIndex: 0 });
    expect(tableChange.properties.map(({ key }) => key)).toEqual(
      Object.keys(TARGET_TABLE_PROPERTIES).toSorted(),
    );
    expect(rowChange.properties.map(({ key }) => key)).toEqual(
      Object.keys(TARGET_ROW_PROPERTIES).toSorted(),
    );
    expect(cellChange.properties.map(({ key }) => key)).toEqual(
      Object.keys(TARGET_CELL_PROPERTIES).toSorted(),
    );

    for (const occurrence of occurrences) {
      expect(Object.hasOwn(occurrence.change, "owner")).toBe(false);
      expectRecursivelyFrozen(occurrence.change);
      expectRecursivelyFrozen(occurrence.owner);
    }

    const { state } = applyProgram(base, result.program);
    expect(
      projectTableGeometry(folioStoryTables(resolveAllChangesInHeadlessState(state, "accept").doc)),
    ).toEqual(projectTableGeometry(folioStoryTables(target)));
    expect(
      projectTableGeometry(folioStoryTables(resolveAllChangesInHeadlessState(state, "reject").doc)),
    ).toEqual(projectTableGeometry(folioStoryTables(base)));
  });

  test("normalizes absent and false presence properties before planning or reporting", () => {
    const base = documentWithFormatting({
      tableFormatting: {},
      rowFormatting: {},
      cellFormatting: {},
    });
    const target = documentWithFormatting({
      tableFormatting: { bidi: false },
      rowFormatting: { header: false, cantSplit: false, hidden: false },
      cellFormatting: { fitText: false, noWrap: false, hideMark: false },
    });

    const result = preflight(base, target);

    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    expect(result.program.type).toBe("unchanged");
    expect(tableGeometryProgramSemanticChangeOccurrences(result.program)).toEqual([]);
    expect(projectTableGeometry(folioStoryTables(base))).toEqual(
      projectTableGeometry(folioStoryTables(target)),
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
      projectTableGeometry(folioStoryTables(resolveAllChangesInHeadlessState(state, "accept").doc)),
    ).toEqual(projectTableGeometry(folioStoryTables(target)));
    expect(
      projectTableGeometry(folioStoryTables(resolveAllChangesInHeadlessState(state, "reject").doc)),
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
    const base = documentWith(table([row([cell("a"), cell("b")])]), table([row([cell("c")])]));
    const target = documentWith(table([row([cell("a"), cell("b")])]), table([row([cell("c")])]));
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

  test("reconstructs a directly authored no-wrap removal", () => {
    const base = documentWith(
      table([
        row([
          cell("same", {
            noWrap: true,
          }),
        ]),
      ]),
    );
    const target = documentWith(table([row([cell("same")])]));

    const result = preflight(base, target);

    expect(projectTableGeometry(folioStoryTables(base))).not.toEqual(
      projectTableGeometry(folioStoryTables(target)),
    );
    expect(result.status).toBe("ready");
    if (result.status === "unsupported") return;
    expect(tableGeometryProgramSemanticChangeOccurrences(result.program).at(0)?.change).toEqual({
      scope: "cell",
      base: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
      target: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
      properties: [
        {
          key: "noWrap",
          base: { type: "present", value: true },
          revised: { type: "absent" },
        },
      ],
    });
    const { state } = applyProgram(base, result.program);
    expect(
      projectTableGeometry(folioStoryTables(resolveAllChangesInHeadlessState(state, "accept").doc)),
    ).toEqual(projectTableGeometry(folioStoryTables(target)));
    expect(
      projectTableGeometry(folioStoryTables(resolveAllChangesInHeadlessState(state, "reject").doc)),
    ).toEqual(projectTableGeometry(folioStoryTables(base)));
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

  test("defers a paired grid change as an explicit structural obligation", () => {
    const base = documentWith(table([row([cell("same")])], { columnWidths: [2400] }));
    const target = documentWith(table([row([cell("same")])], { columnWidths: [3600] }));
    const owner = pairing();

    const component = {};
    const components = preflightTableGeometryComponents({
      baseTables: folioStoryTables(base),
      targetTables: targetTablesOf(target),
      components: [{ component, pairings: [owner] }],
      tableGridChanges: "defer-to-table-structure",
    });

    expect(components.status).toBe("ready");
    if (components.status === "unsupported") return;
    const result = components.components.at(0)?.result;
    expect(result?.status).toBe("ready");
    if (!result || result.status === "unsupported") return;
    expect(tableGeometryProgramSemanticChangeOccurrences(result.program)).toEqual([]);
    expect(tableGeometryProgramTableGridTransitions(result.program)).toEqual([owner]);
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

  test("shares geometry limits across independent components", () => {
    const base = documentWith(
      table([row([cell("first")])], { width: 2_000, widthType: "dxa" }),
      table([row([cell("second")])], { width: 3_000, widthType: "dxa" }),
    );
    const target = documentWith(
      table([row([cell("first")])], { width: 2_400, widthType: "dxa" }),
      table([row([cell("second")])], { width: 3_600, widthType: "dxa" }),
    );
    const first = Object.freeze({ table: "first" });
    const second = Object.freeze({ table: "second" });

    const result = preflightTableGeometryComponents({
      baseTables: folioStoryTables(base),
      targetTables: targetTablesOf(target),
      components: [
        { component: first, pairings: [pairing(coordinate(0), coordinate(0))] },
        { component: second, pairings: [pairing(coordinate(1), coordinate(1))] },
      ],
      limits: { ...DEFAULT_TABLE_GEOMETRY_PREFLIGHT_LIMITS, maxChanges: 1 },
    });

    expect(result).toEqual({
      status: "unsupported",
      issue: {
        reason: "limit-exceeded",
        limit: "maxChanges",
        maximum: 1,
        actual: 2,
      },
    });
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
      table([row([staleCell], { height: 320, heightRule: "exact" })], {
        width: 4800,
        widthType: "dxa",
        justification: "left",
        columnWidths: [4800],
      }),
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

    expect(execution).toMatchObject({
      status: "unsupported",
      issue: { reason: "missing-live-node" },
    });
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
    const reversed = [pairing(coordinate(0), coordinate(1)), pairing(coordinate(1), coordinate(0))];
    const forward = [...reversed].reverse();
    const first = preflight(base, target, reversed);
    const second = preflight(base, target, forward);
    expect(first.status).toBe("ready");
    expect(second.status).toBe("ready");
    if (first.status === "unsupported" || second.status === "unsupported") return;

    const firstReceipt = applyProgram(base, first.program).receipt;
    const secondReceipt = applyProgram(base, second.program).receipt;
    expect(firstReceipt).toEqual(secondReceipt);
    expect(
      firstReceipt.revisions.map(({ base: source, target: destination }) => ({
        base: source.tableIndex,
        target: destination.tableIndex,
      })),
    ).toEqual([
      { base: 1, target: 0 },
      { base: 0, target: 1 },
    ]);
  });
});
