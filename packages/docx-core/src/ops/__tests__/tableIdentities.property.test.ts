import { expect, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type {
  Document,
  DocumentBody,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { IDENTITY_VALIDATION_REASONS, type DocumentIdLease } from "../../model/identity";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { paragraphLength } from "../offsets";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  SPLIT_HALVES,
  type DocumentOp,
  type SplitHalf,
} from "../types";

const lease = {
  paragraph: { start: 1, end: 65535 },
  revision: { start: 1, end: 1024 },
  comment: { start: 1025, end: 2048 },
  table: { prefix: "tables", start: 0, end: 65535 },
  row: { prefix: "rows", start: 0, end: 65535 },
  cell: { prefix: "cells", start: 0, end: 65535 },
} satisfies DocumentIdLease;

const paragraph = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: "abcd" }] }],
});

// The caller owns this monotonic cursor, including ids removed by subsequent edits.
const mint = () => {
  let cursor = 0;
  const issued = new Set<string>();
  const next = (kind: "paragraph" | "row" | "cell"): string => {
    cursor++;
    let id = cursor.toString(16).padStart(8, "0").toUpperCase();
    if (kind !== "paragraph") {
      const digits: string[] = [];
      for (let remaining = cursor; remaining > 0; remaining = Math.floor(remaining / 32)) {
        const digit = "abcdefghijklmnopqrstuvwxyz234567".at(remaining % 32);
        if (digit === undefined) panic("Base32 digit is present.");
        digits.unshift(digit);
      }
      const prefix = kind === "row" ? "r_rows_" : "c_cells_";
      id = `${prefix}${digits.join("")}`;
    }
    expect(issued.has(id)).toBe(false);
    issued.add(id);
    return id;
  };
  return next;
};
type Mint = ReturnType<typeof mint>;
const cell = (next: Mint): TableCell => ({
  type: "tableCell",
  cellId: next("cell"),
  content: [paragraph(next("paragraph"))],
});
const row = (next: Mint, columns: number): TableRow => ({
  type: "tableRow",
  rowId: next("row"),
  cells: Array.from({ length: columns }, () => cell(next)),
});
const fixture = (next: Mint): Document => ({
  package: {
    document: {
      content: [
        paragraph(next("paragraph")),
        {
          type: "table",
          tableId: "t_tables_a",
          columnWidths: [900, 900],
          rows: [row(next, 2), row(next, 2)],
        },
      ],
    },
  },
});

// Deliberately independent of the production identity enumerator and sidecar builder.
const identities = (document: Document): string[] => {
  const out: string[] = [];
  const visit = (blocks: DocumentBody["content"]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        if (block.paraId === undefined) panic("Every generated paragraph has an id.");
        out.push(block.paraId.toUpperCase());
      } else if (block.type === "table") {
        if (block.tableId === undefined) panic("Every generated table has an id.");
        out.push(block.tableId);
        for (const value of block.rows) {
          if (value.rowId === undefined) panic("Every generated row has an id.");
          out.push(value.rowId);
          for (const entry of value.cells) {
            if (entry.cellId === undefined) panic("Every generated cell has an id.");
            out.push(entry.cellId);
            visit(entry.content);
          }
        }
      }
    }
  };
  visit(document.package.document.content);
  return out.sort();
};
const tableOf = (document: Document): Table => {
  const value = document.package.document.content.find((block) => block.type === "table");
  if (value === undefined) panic("The generated table remains present.");
  return value;
};
const anchor = (value: TableCell): string => {
  const block = value.content.find((entry) => entry.type === "paragraph");
  if (block?.paraId === undefined) panic("The generated cell has a paragraph anchor.");
  return block.paraId;
};

const actions = [
  "insertRow",
  "deleteRow",
  "insertCell",
  "deleteCell",
  "splitBlock",
  "joinBlock",
  "mergeCells",
  "splitCell",
] as const;
type Action = (typeof actions)[number];
const halfChoices = [undefined, ...Object.values(SPLIT_HALVES)];
type Step = {
  action: Action;
  choice: number;
  splitHalf?: SplitHalf | undefined;
  survivor?: SplitHalf | undefined;
};
type BuildOpOptions = Step & { document: Document; next: Mint };
const buildOp = ({
  document,
  action,
  choice,
  splitHalf,
  survivor,
  next,
}: BuildOpOptions): DocumentOp => {
  const table = tableOf(document);
  const firstRow = table.rows.at(0);
  const firstCell = firstRow?.cells.at(0);
  if (firstRow === undefined || firstCell === undefined) panic("The generated table is nonempty.");
  const target = { story: OP_STORIES.MAIN, blockId: anchor(firstCell) };
  const widths = table.columnWidths ?? [];
  switch (action) {
    case "insertRow":
      return {
        ...target,
        type: DOCUMENT_OP_TYPES.INSERT_ROW,
        idLease: lease,
        at: choice % (table.rows.length + 1),
        row: row(next, widths.length),
      };
    case "deleteRow": {
      if (table.rows.length === 1) return buildOp({ document, action: "insertRow", choice, next });
      const selected = table.rows.at(choice % table.rows.length)?.cells.at(0);
      if (selected === undefined) panic("Selected row has a cell.");
      return {
        type: DOCUMENT_OP_TYPES.DELETE_ROW,
        story: OP_STORIES.MAIN,
        blockId: anchor(selected),
      };
    }
    case "insertCell":
      return {
        ...target,
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        idLease: lease,
        column: choice % (widths.length + 1),
        width: 900,
        newCellIds: table.rows.map(() => next("cell")),
        newBlockIds: table.rows.map(() => next("paragraph")),
      };
    case "deleteCell":
      if (widths.length === 1) return buildOp({ document, action: "insertCell", choice, next });
      return { ...target, type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: choice % widths.length };
    case "splitBlock": {
      const block = document.package.document.content.at(0);
      if (block?.type !== "paragraph" || block.paraId === undefined)
        panic("The leading paragraph exists.");
      return {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: {
          story: OP_STORIES.MAIN,
          blockId: block.paraId,
          offset: choice % (paragraphLength(block) + 1),
        },
        idLease: lease,
        newBlockId: next("paragraph"),
        ...(splitHalf === undefined ? {} : { newHalf: splitHalf }),
      };
    }
    case "joinBlock": {
      const first = document.package.document.content.at(0);
      const second = document.package.document.content.at(1);
      if (second?.type !== "paragraph")
        return buildOp({ document, action: "splitBlock", choice, splitHalf, next });
      if (first?.type !== "paragraph" || first.paraId === undefined || second.paraId === undefined)
        panic("Joined paragraphs have ids.");
      return {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: first.paraId,
        nextBlockId: second.paraId,
        ...(survivor === undefined ? {} : { survivor }),
      };
    }
    case "mergeCells":
      if (firstRow.cells.length < 2)
        return buildOp({ document, action: "insertCell", choice, next });
      return {
        ...target,
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        idLease: lease,
        top: 0,
        bottom: 1,
        left: 0,
        right: 2,
        newCellIds: [],
        newBlockIds: [],
      };
    case "splitCell": {
      const span = firstCell.formatting?.gridSpan ?? 1;
      if (span === 1) return buildOp({ document, action: "mergeCells", choice, next });
      return {
        ...target,
        type: DOCUMENT_OP_TYPES.SPLIT_CELL,
        idLease: lease,
        newCellIds: Array.from({ length: span - 1 }, () => next("cell")),
        newBlockIds: Array.from({ length: span - 1 }, () => next("paragraph")),
      };
    }
  }
};

const runSequence = (steps: readonly Step[]): void => {
  const next = mint();
  const original = fixture(next);
  let document = original;
  const retired = new Set<string>();
  const inverses: (readonly DocumentOp[])[] = [];
  for (const step of steps) {
    // Normalize spans before geometry edits; check normalization as an ordinary edit.
    const table = tableOf(document);
    const spanned = table.rows.at(0)?.cells.at(0)?.formatting?.gridSpan ?? 1;
    const pending: Step[] = [];
    if (spanned > 1 && step.action !== "splitCell")
      pending.push({ action: "splitCell", choice: 0 });
    pending.push(step);
    for (const current of pending) {
      const before = identities(document);
      const op = buildOp({ document, ...current, next });
      const edit = applyDocumentOp(document, op)._unsafeUnwrap();
      const after = identities(edit.document);
      expect(new Set(after).size).toBe(after.length);
      const previous = new Set(before);
      const present = new Set(after);
      for (const id of after) if (!previous.has(id)) expect(retired.has(id)).toBe(false);
      for (const id of before) if (!present.has(id)) retired.add(id);
      if (op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) {
        expect(edit.document.package.document.content.at(0)?.type).toBe("paragraph");
        const first = edit.document.package.document.content.at(0);
        if (first?.type !== "paragraph") panic("Split starts with a paragraph.");
        const second = edit.document.package.document.content.at(1);
        if (second?.type !== "paragraph") panic("Split creates a second paragraph.");
        const newHalf = op.newHalf ?? SPLIT_HALVES.SECOND;
        expect(first.paraId).toBe(newHalf === SPLIT_HALVES.FIRST ? op.newBlockId : op.at.blockId);
        expect(second.paraId).toBe(newHalf === SPLIT_HALVES.SECOND ? op.newBlockId : op.at.blockId);
        expect(after.filter((id) => !previous.has(id))).toEqual([op.newBlockId]);
        expect(before.filter((id) => !present.has(id))).toEqual([]);
        const newId = Number.parseInt(op.newBlockId, 16);
        expect(newId).toBeGreaterThanOrEqual(op.idLease.paragraph.start);
        expect(newId).toBeLessThanOrEqual(op.idLease.paragraph.end);
      }
      if (op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS) {
        const first = edit.document.package.document.content.at(0);
        if (first?.type !== "paragraph") panic("Join leaves a paragraph.");
        const survivorSide = op.survivor ?? SPLIT_HALVES.SECOND;
        expect(first.paraId).toBe(
          survivorSide === SPLIT_HALVES.FIRST ? op.blockId : op.nextBlockId,
        );
        const retiredId = survivorSide === SPLIT_HALVES.FIRST ? op.nextBlockId : op.blockId;
        expect(before.filter((id) => !present.has(id))).toEqual([retiredId]);
        expect(after.filter((id) => !previous.has(id))).toEqual([]);
      }
      const restored = applyDocumentOps(edit.document, edit.inverse)._unsafeUnwrap().document;
      expect(identities(restored)).toEqual(before);
      expect(restored).toStrictEqual(document);
      inverses.push(edit.inverse);
      document = edit.document;
    }
  }
  const restored = applyDocumentOps(document, inverses.reverse().flat())._unsafeUnwrap().document;
  expect(identities(restored)).toEqual(identities(original));
  expect(restored).toStrictEqual(original);
};

const families = {
  rows: ["insertRow", "deleteRow"],
  cells: ["insertCell", "deleteCell", "mergeCells", "splitCell"],
  blocks: ["splitBlock", "joinBlock"],
  mixed: actions,
} as const satisfies Record<string, readonly Action[]>;
for (const [family, members] of Object.entries(families)) {
  test(
    `leased ${family} sequences preserve unique identities and exact inverses`,
    () => {
      assertProperty(
        fc.property(
          fc.array(
            fc.record({
              action: fc.constantFrom(...members),
              choice: fc.nat({ max: 255 }),
              splitHalf: fc.constantFrom(...halfChoices),
              survivor: fc.constantFrom(...halfChoices),
            }),
            { minLength: 1, maxLength: 24 },
          ),
          (steps) => {
            const prefix: Step[] = [];
            if (family === "blocks" || family === "mixed") {
              // Exercise every default/explicit split and join choice on every run.
              for (const splitHalf of halfChoices) {
                for (const survivor of halfChoices) {
                  prefix.push(
                    { action: "splitBlock", choice: 2, splitHalf },
                    { action: "joinBlock", choice: 0, survivor },
                  );
                }
              }
              // Pin interior and end cuts as well as the start cut in the family prelude.
              prefix.push(
                { action: "splitBlock", choice: 2 },
                { action: "joinBlock", choice: 1 },
                { action: "splitBlock", choice: 4 },
                { action: "joinBlock", choice: 0 },
              );
            }
            runSequence([...prefix, ...members.map((action) => ({ action, choice: 0 })), ...steps]);
          },
        ),
        { id: `leased-${family}-identity-sequences`, numRuns: 50 },
      );
    },
    propertyTestTimeout(30_000),
  );
}

test("row insertion refuses invalid caller identities without mutation", () => {
  const next = mint();
  const source = fixture(next);
  const op = buildOp({ document: source, action: "insertRow", choice: 0, next });
  if (op.type !== DOCUMENT_OP_TYPES.INSERT_ROW) panic("Fixture builds row insertion.");
  const existing = tableOf(source).rows.at(0)?.rowId;
  if (existing === undefined) panic("Existing row has an id.");
  const cases = [
    {
      op: { ...op, row: { ...op.row, rowId: "r_other_a" } },
      reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
    },
    {
      op: { ...op, row: { ...op.row, rowId: existing } },
      reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
    },
    {
      op: { ...op, idLease: { ...lease, row: { ...lease.row, end: -1 } } },
      reason: IDENTITY_VALIDATION_REASONS.INVALID_LEASE,
    },
  ];
  for (const entry of cases) {
    const before = structuredClone(source);
    expect(applyDocumentOp(source, entry.op)._unsafeUnwrapErr().reason).toBe(entry.reason);
    expect(source).toStrictEqual(before);
  }
});

test("column and block creation validate caller identities without mutation", () => {
  const next = mint();
  const source = fixture(next);
  const column = buildOp({ document: source, action: "insertCell", choice: 0, next });
  const split = buildOp({ document: source, action: "splitBlock", choice: 0, next });
  if (
    column.type !== DOCUMENT_OP_TYPES.INSERT_COLUMN ||
    split.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK
  )
    panic("Fixtures build column insertion and block splitting.");
  const existingCell = tableOf(source).rows.at(0)?.cells.at(0)?.cellId;
  const existingBlock = source.package.document.content.at(0);
  if (
    existingCell === undefined ||
    existingBlock?.type !== "paragraph" ||
    existingBlock.paraId === undefined
  )
    panic("Existing fixture identities are present.");
  const cases = [
    {
      op: { ...column, newCellIds: ["c_other_a", next("cell")] },
      reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
    },
    {
      op: { ...column, newCellIds: [existingCell, next("cell")] },
      reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
    },
    {
      op: { ...column, newBlockIds: ["00010000", next("paragraph")] },
      reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
    },
    {
      op: { ...column, idLease: { ...lease, cell: { ...lease.cell, end: -1 } } },
      reason: IDENTITY_VALIDATION_REASONS.INVALID_LEASE,
    },
    {
      op: { ...split, newBlockId: "00010000" },
      reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
    },
    {
      op: { ...split, newBlockId: existingBlock.paraId },
      reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
    },
    {
      op: { ...split, idLease: { ...lease, paragraph: { start: 2, end: 1 } } },
      reason: IDENTITY_VALIDATION_REASONS.INVALID_LEASE,
    },
  ];
  for (const entry of cases) {
    const before = structuredClone(source);
    expect(applyDocumentOp(source, entry.op)._unsafeUnwrapErr().reason).toBe(entry.reason);
    expect(source).toStrictEqual(before);
  }
});

const extraStories = ["header", "footer", "footnote", "endnote", "nested"] as const;
for (const story of extraStories) {
  test(`row ids cannot collide with identities in a ${story} story or container`, () => {
    const next = mint();
    const source = fixture(next);
    const externalRow = row(next, 1);
    const external: Table = { type: "table", tableId: "t_tables_b", rows: [externalRow] };
    switch (story) {
      case "header":
        source.package.headers = new Map([
          ["rId1", { type: "header", hdrFtrType: "default", content: [external] }],
        ]);
        break;
      case "footer":
        source.package.footers = new Map([
          ["rId1", { type: "footer", hdrFtrType: "default", content: [external] }],
        ]);
        break;
      case "footnote":
        source.package.footnotes = [{ type: "footnote", id: 1, content: [external] }];
        break;
      case "endnote":
        source.package.endnotes = [{ type: "endnote", id: 1, content: [external] }];
        break;
      case "nested": {
        const owner = tableOf(source).rows.at(0)?.cells.at(0);
        if (owner === undefined) panic("Nested table owner exists.");
        owner.content.unshift(external);
        break;
      }
    }
    const op = buildOp({ document: source, action: "insertRow", choice: 0, next });
    if (op.type !== DOCUMENT_OP_TYPES.INSERT_ROW) panic("Fixture builds row insertion.");
    const before = structuredClone(source);
    const collision = { ...op, row: { ...op.row, rowId: externalRow.rowId } };
    expect(applyDocumentOp(source, collision)._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.ID_COLLISION,
    );
    expect(source).toStrictEqual(before);
  });
}
