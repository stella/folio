/**
 * A direct cell merge across rows with different pending insertions and
 * deletions keeps each row's change on the words it gives the merged cell:
 * accepting or rejecting every change afterwards gives the merge of the rows
 * that decision keeps, and the inverse restores the table exactly. Where the
 * merged cell ends with words accepting removes, its last paragraph takes the
 * words and properties of the one before, as deleting the paragraphs at a
 * cell's end does, and keeps its own id.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphFormatting,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type MergeCellsOp,
  OP_STORIES,
} from "../types";

const INFO = { author: "Reviewer", date: "2026-01-01T00:00:00Z" };
const ROWS = 3;

/** The paragraph id of the cell in `row` and `column`. */
const at = (row: number, column: number) => `000${row}${column}0A0`;
const FRESH = Array.from({ length: 4 }, (_, index) => `00F00${index}A0`);
const NEW_REVISION_IDS = Array.from({ length: 40 }, (_, index) => 500 + index);

const ROW_FATES = ["stands", "inserted", "deleted"] as const;
type RowFate = (typeof ROW_FATES)[number];

/** Whether a row's words already record its change, as a tracked row's usually do. */
const CONTENT_KINDS = ["plain", "marked"] as const;
type ContentKind = (typeof CONTENT_KINDS)[number];

/** Each row's paragraph states other properties, so a join shows whose it kept. */
const FORMATTING: readonly (ParagraphFormatting | undefined)[] = [
  undefined,
  { alignment: "center" },
  { alignment: "right" },
];

type ShapeOptions = { fates: readonly RowFate[]; content: ContentKind };

const paragraphOf = (
  row: number,
  column: number,
  { fates, content }: ShapeOptions,
): Paragraph => {
  const fate = fates[row];
  const paraId = at(row, column);
  const run = {
    type: "run" as const,
    content: [{ type: "text" as const, text: paraId }],
  };
  const formatting = FORMATTING[row];
  const base = {
    type: "paragraph" as const,
    paraId,
    ...(formatting === undefined ? {} : { formatting }),
  };
  if (content === "plain" || fate === "stands")
    return { ...base, content: [run] };
  const id = 100 + row * 10 + column * 2;
  const kind = fate === "inserted" ? "insertion" : "deletion";
  return {
    ...base,
    content: [{ type: kind, info: { id, ...INFO }, content: [run] }],
    pPrMark: {
      kind: fate === "inserted" ? "ins" : "del",
      info: { id: id + 1, ...INFO },
    },
  };
};

const tableOf = (options: ShapeOptions): Table => ({
  type: "table",
  columnWidths: [2000, 2000],
  rows: Array.from({ length: ROWS }, (_, row): TableRow => {
    const fate = options.fates[row];
    return {
      type: "tableRow",
      cells: [0, 1].map((column) => ({
        type: "tableCell" as const,
        content: [paragraphOf(row, column, options)],
      })),
      ...(fate === "stands"
        ? {}
        : {
            structuralChange: {
              type:
                fate === "inserted" ? "tableRowInsertion" : "tableRowDeletion",
              info: { id: 10 + row, ...INFO },
            },
          }),
    };
  }),
});

const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [table, { type: "paragraph", paraId: "00000001", content: [] }],
    },
  },
});
const firstTable = (document: Document): Table => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new Error("The table stays first.");
  return table;
};

const firstParagraphId = (table: Table): string => {
  const paragraph = table.rows.at(0)?.cells.at(0)?.content.at(0);
  if (paragraph?.type !== "paragraph" || paragraph.paraId === undefined)
    throw new Error("The table starts with a paragraph.");
  return paragraph.paraId;
};

/** Every revision id the document's records hold. */
const revisionIds = (node: unknown): number[] => {
  if (Array.isArray(node)) return node.flatMap(revisionIds);
  if (node === null || typeof node !== "object") return [];
  const info: unknown = Reflect.get(node, "info");
  const own =
    info !== null && typeof info === "object"
      ? [Reflect.get(info, "id")].filter((id) => typeof id === "number")
      : [];
  return [
    ...own,
    ...Object.entries(node)
      .filter(([key]) => key !== "info")
      .flatMap(([, value]) => revisionIds(value)),
  ];
};

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const DECISIONS = ["accept", "reject"] as const;
type Decision = (typeof DECISIONS)[number];
const resolveAll = (document: Document, decision: Decision) => {
  const ids = [...new Set(revisionIds(document))];
  if (ids.length === 0) return document;
  return apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision,
  }).document;
};

/** The fate a decision removes a row by. */
const REMOVED_BY = {
  accept: "deleted",
  reject: "inserted",
} as const satisfies Record<Decision, RowFate>;

const MERGES = {
  /** The first column of every row: each row keeps its second cell. */
  down: { left: 0, right: 1 },
  /** Every cell: the rows after the first keep nothing of their own. */
  whole: { left: 0, right: 2 },
} as const;
type MergeName = keyof typeof MERGES;

type MergeOpOptions = { name: MergeName; rows: number; blockId: string };
const mergeOp = ({ name, rows, blockId }: MergeOpOptions): MergeCellsOp => ({
  type: DOCUMENT_OP_TYPES.MERGE_CELLS,
  story: OP_STORIES.MAIN,
  blockId,
  top: 0,
  bottom: rows,
  ...MERGES[name],
  newBlockIds: FRESH.slice(0, rows - 1),
  newIds: { revision: NEW_REVISION_IDS },
});

/** The body with every new paragraph's id read as one. */
const withoutFreshIds = (document: Document): unknown =>
  JSON.parse(
    JSON.stringify(document.package.document.content).replaceAll(
      /"00F00\dA0"/g,
      `"new"`,
    ),
  );

type Expected = { type: "applies" } | { type: "refused"; message: string };

/**
 * Where the merged cell ends with words one decision removes, the paragraph
 * that stands before them takes the cell's final mark, which a paragraph
 * whose own mark is under another row's change cannot.
 */
const expectedOf = (name: MergeName, fates: readonly RowFate[]): Expected => {
  const last = fates.at(-1);
  const before = fates.findLast((fate) => fate !== last);
  if (last !== "stands" && before !== undefined && before !== "stands")
    return {
      type: "refused",
      message:
        "Resolve the pending paragraph changes at the end of the merged cell before merging.",
    };
  // A first row under a change of its own would take the rows it takes in
  // whole with it.
  if (name === "whole" && fates[0] !== "stands")
    return {
      type: "refused",
      message:
        "Resolve the pending insertion or deletion of the merge's first row before merging rows it would take whole.",
    };
  return { type: "applies" };
};

/** The cell's last paragraph with another id. */
const renamedLast = (cell: TableCell, paraId: string): TableCell => {
  const last = cell.content.at(-1);
  if (last?.type !== "paragraph")
    throw new Error("The merged cell ends with a paragraph.");
  return {
    ...cell,
    content: [...cell.content.slice(0, -1), { ...last, paraId }],
  };
};

/**
 * The merge of the rows a decision keeps, in the document that decision
 * leaves. Accepting a merged cell that ends with deleted words leaves its
 * last paragraph's id on the words before them.
 */
const referenceOf = (
  original: Document,
  name: MergeName,
  fates: readonly RowFate[],
  decision: Decision,
): Document => {
  const resolved = resolveAll(original, decision);
  const kept = fates.filter((fate) => fate !== REMOVED_BY[decision]).length;
  if (kept === 0) return resolved;
  const merged =
    kept > 1 || name === "whole"
      ? apply(
          resolved,
          mergeOp({
            name,
            rows: kept,
            blockId: firstParagraphId(firstTable(resolved)),
          }),
        ).document
      : resolved;
  if (decision !== "accept" || fates.at(-1) !== "deleted") return merged;
  const table = firstTable(merged);
  const [first, ...rest] = table.rows;
  const [cell, ...others] = first?.cells ?? [];
  if (first === undefined || cell === undefined)
    throw new Error("The merged table keeps its first cell.");
  const renamed: Table = {
    ...table,
    rows: [
      {
        ...first,
        cells: [
          renamedLast(cell, at(fates.length - 1, MERGES[name].right - 1)),
          ...others,
        ],
      },
      ...rest,
    ],
  };
  return {
    package: {
      document: {
        content: [renamed, ...merged.package.document.content.slice(1)],
      },
    },
  };
};

const patterns = (): RowFate[][] => {
  const out: RowFate[][] = [];
  const visit = (prefix: RowFate[]) => {
    if (prefix.length === ROWS) {
      if (prefix.some((fate) => fate !== "stands")) out.push(prefix);
      return;
    }
    for (const fate of ROW_FATES) visit([...prefix, fate]);
  };
  visit([]);
  return out;
};

describe.each(Object.keys(MERGES) as MergeName[])(
  "a direct %s merge",
  (name) => {
    describe.each(CONTENT_KINDS)("of %s row words", (content) => {
      test.each(patterns().map((fates) => [fates.join(" "), fates] as const))(
        "across rows that %s",
        (_, fates) => {
          const original = documentOf(tableOf({ fates, content }));
          const snapshot = structuredClone(original);
          const result = applyDocumentOp(
            original,
            mergeOp({ name, rows: ROWS, blockId: at(0, 0) }),
          );
          expect(original).toStrictEqual(snapshot);
          const expected = expectedOf(name, fates);
          if (expected.type === "refused") {
            expect(result.isErr() && result.error.reason).toBe(
              DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
            );
            expect(result.isErr() && result.error.message).toBe(
              expected.message,
            );
            return;
          }
          if (result.isErr()) throw result.error;
          const edited = result.value.document;
          for (const decision of DECISIONS) {
            const resolved = resolveAll(edited, decision);
            expect(revisionIds(resolved)).toEqual([]);
            expect(withoutFreshIds(resolved)).toEqual(
              withoutFreshIds(referenceOf(original, name, fates, decision)),
            );
          }
          let undone = edited;
          for (const inverse of result.value.inverse)
            undone = apply(undone, inverse).document;
          expect(undone).toStrictEqual(original);
        },
      );
    });
  },
);
