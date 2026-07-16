/**
 * Regression: foreign editors (MS Word Online, in particular) emit a
 * fresh `w:id` per atomic edit even when the edits share an author and
 * timestamp. The extractor must coalesce by (author, date) so a single
 * logical revision burst surfaces as ONE sidebar card, and the dropped
 * ids must be tucked into `coalescedRevisionIds` so Accept/Reject still
 * clear every site in one click.
 */

import { describe, test, expect } from "bun:test";
import { Schema } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { extractTrackedChanges } from "./extractTrackedChanges";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { pPrIns: { default: null }, pPrDel: { default: null } },
      toDOM: () => ["p", 0],
    },
    tableRow: {
      content: "tableCell+",
      attrs: { trIns: { default: null }, trDel: { default: null } },
      toDOM: () => ["tr", 0],
    },
    tableCell: {
      content: "paragraph+",
      attrs: { cellMarker: { default: null } },
      toDOM: () => ["td", 0],
    },
    table: { content: "tableRow+", group: "block", toDOM: () => ["table", 0] },
    text: { group: "inline" },
  },
  marks: {
    insertion: {
      attrs: { revisionId: { default: 0 }, author: { default: "" }, date: { default: null } },
      toDOM: () => ["ins", 0],
    },
    deletion: {
      attrs: { revisionId: { default: 0 }, author: { default: "" }, date: { default: null } },
      toDOM: () => ["del", 0],
    },
    comment: {
      attrs: { commentId: { default: 0 } },
      toDOM: () => ["span", 0],
    },
  },
});

const AUTHOR = "Docx Editor User 960";
const DATE = "2026-05-28T20:28:35.944Z";

function makeState(doc: ReturnType<typeof schema.node>): EditorState {
  return EditorState.create({ doc });
}

describe("extractTrackedChanges: foreign-doc coalescing by (author, date)", () => {
  test("5 insertions with distinct w:ids but same (author, date) collapse to ONE card", () => {
    const ins = (id: number, text: string) =>
      schema.text(text, [
        schema.marks.insertion.create({ revisionId: id, author: AUTHOR, date: DATE }),
      ]);
    const doc = schema.nodes.doc.create({}, [
      schema.nodes.paragraph.create({}, [ins(1388975360, "fdsfsd")]),
      schema.nodes.paragraph.create({}, [ins(47262383, "fdsfsdf")]),
      schema.nodes.paragraph.create({}, [ins(1323221525, "dsfsd")]),
      schema.nodes.paragraph.create({}, [ins(737865714, "fds")]),
      schema.nodes.paragraph.create({}, [ins(186027604, "fsd")]),
    ]);
    const { entries } = extractTrackedChanges(makeState(doc));
    expect(entries).toHaveLength(1);
    const e = entries[0]!;
    expect(e.type).toBe("insertion");
    expect(e.author).toBe(AUTHOR);
    expect(e.date).toBe(DATE);
    // Coalesced ids cover the 4 absorbed entries (the primary lives on `revisionId`).
    expect(new Set([e.revisionId, ...(e.coalescedRevisionIds ?? [])])).toEqual(
      new Set([1388975360, 47262383, 1323221525, 737865714, 186027604]),
    );
  });

  test("paragraph-mark insertions with distinct ids but same (author, date) hide behind one inline card", () => {
    const ins = (id: number, text: string) =>
      schema.text(text, [
        schema.marks.insertion.create({ revisionId: id, author: AUTHOR, date: DATE }),
      ]);
    const pPrIns = (id: number) => ({ revisionId: id, author: AUTHOR, date: DATE });
    const doc = schema.nodes.doc.create({}, [
      schema.nodes.paragraph.create({ pPrIns: pPrIns(1694997150) }, [ins(1388975360, "fdsfsd")]),
      schema.nodes.paragraph.create({ pPrIns: pPrIns(1254058768) }, [ins(47262383, "fdsfsdf")]),
    ]);
    const { entries } = extractTrackedChanges(makeState(doc));
    expect(entries).toHaveLength(1);
    expect(entries[0]!.type).toBe("insertion");
  });

  test('whole-table insert: rows with distinct trIns ids but same (author, date) surface "Inserted table"', () => {
    const trIns = (id: number) => ({ revisionId: id, author: AUTHOR, date: DATE });
    const emptyCell = () =>
      schema.nodes.tableCell.create({}, [schema.nodes.paragraph.create({}, [])]);
    const row = (id: number) =>
      schema.nodes.tableRow.create({ trIns: trIns(id) }, [emptyCell(), emptyCell()]);
    const table = schema.nodes.table.create({}, [row(844706625), row(694611694)]);
    const doc = schema.nodes.doc.create({}, [table]);
    const { entries } = extractTrackedChanges(makeState(doc));
    const tableEntry = entries.find((e) => e.type === "tableInserted");
    expect(tableEntry).toBeTruthy();
    // The two row ids should both be reachable from the card.
    const allIds = new Set([tableEntry!.revisionId, ...(tableEntry!.coalescedRevisionIds ?? [])]);
    expect(allIds).toEqual(new Set([844706625, 694611694]));
  });

  test("table-cell insertion and deletion markers surface as structural changes", () => {
    const cell = (kind: "ins" | "del", revisionId: number, text: string, date = DATE) =>
      schema.nodes.tableCell.create(
        {
          cellMarker: {
            kind,
            info: { revisionId, author: AUTHOR, date },
          },
        },
        [schema.nodes.paragraph.create({}, [schema.text(text)])],
      );
    const doc = schema.nodes.doc.create({}, [
      schema.nodes.table.create({}, [
        schema.nodes.tableRow.create({}, [cell("ins", 101, "Inserted")]),
        schema.nodes.tableRow.create({}, [cell("del", 102, "Deleted", "2026-05-28T20:30:00.000Z")]),
      ]),
    ]);

    const { entries } = extractTrackedChanges(makeState(doc));

    expect(entries.map(({ type, revisionId, text }) => ({ type, revisionId, text }))).toEqual([
      { type: "cellInserted", revisionId: 101, text: "Inserted" },
      { type: "cellDeleted", revisionId: 102, text: "Deleted" },
    ]);
  });

  test("opposing structural changes with the same author and timestamp stay separate", () => {
    const cell = (kind: "ins" | "del", revisionId: number, text: string) =>
      schema.nodes.tableCell.create(
        {
          cellMarker: {
            kind,
            info: { revisionId, author: AUTHOR, date: DATE },
          },
        },
        [schema.nodes.paragraph.create({}, [schema.text(text)])],
      );
    const doc = schema.nodes.doc.create({}, [
      schema.nodes.table.create({}, [
        schema.nodes.tableRow.create({}, [cell("ins", 201, "Inserted")]),
        schema.nodes.tableRow.create({}, [cell("del", 202, "Deleted")]),
      ]),
    ]);

    const { entries } = extractTrackedChanges(makeState(doc));

    expect(entries.map(({ type, revisionId }) => ({ type, revisionId }))).toEqual([
      { type: "cellInserted", revisionId: 201 },
      { type: "cellDeleted", revisionId: 202 },
    ]);
  });

  test("distinct (author, date) bursts stay as separate cards (we are not over-coalescing)", () => {
    const ins = (id: number, author: string, date: string, text: string) =>
      schema.text(text, [schema.marks.insertion.create({ revisionId: id, author, date })]);
    const doc = schema.nodes.doc.create({}, [
      schema.nodes.paragraph.create({}, [ins(1, "Jane", "2026-05-28T20:28:35.944Z", "foo")]),
      schema.nodes.paragraph.create({}, [ins(2, "Jane", "2026-05-28T20:30:00.000Z", "bar")]),
      schema.nodes.paragraph.create({}, [ins(3, "Bob", "2026-05-28T20:28:35.944Z", "baz")]),
    ]);
    const { entries } = extractTrackedChanges(makeState(doc));
    expect(entries).toHaveLength(3);
  });
});
