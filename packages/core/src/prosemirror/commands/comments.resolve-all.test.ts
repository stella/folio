import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { history, undoDepth } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection, type Command, type Transaction } from "prosemirror-state";

import {
  getChangeTrackerState,
  ParagraphChangeTrackerExtension,
} from "../extensions/features/ParagraphChangeTrackerExtension";
import { schema } from "../schema";
import { acceptAllChanges, acceptChange, rejectAllChanges, rejectChange } from "./comments";

const AUTHOR = "Reviewer";
const DATE = "2026-09-09T00:00:00.000Z";
const trackerPlugin = ParagraphChangeTrackerExtension().onSchemaReady({ schema }).plugins?.at(0);
if (!trackerPlugin) {
  throw new Error("Expected the paragraph change tracker plugin.");
}

const revision = (revisionId: number) => ({ revisionId, author: AUTHOR, date: DATE });
const paragraphMark = (kind: "ins" | "del" | "moveTo" | "moveFrom", id: number) => ({
  kind,
  info: { id, author: AUTHOR, date: DATE },
});

type ParagraphCase = {
  mark: "none" | "ins" | "del" | "moveTo" | "moveFrom";
  inline: "plain" | "insertion" | "deletion" | "mixed" | "property";
  paragraphProperty: boolean;
  sectionProperty: boolean;
};

const paragraphCase = fc.record({
  mark: fc.constantFrom("none", "ins", "del", "moveTo", "moveFrom"),
  inline: fc.constantFrom("plain", "insertion", "deletion", "mixed", "property"),
  paragraphProperty: fc.boolean(),
  sectionProperty: fc.boolean(),
}) satisfies fc.Arbitrary<ParagraphCase>;

const paragraph = (item: ParagraphCase, index: number): PMNode => {
  const id = index * 10 + 1;
  const inserted = schema.marks.insertion.create(revision(id));
  const deleted = schema.marks.deletion.create(revision(id + 1));
  const formatting = schema.marks.runPropertyChange.create({
    changes: [
      {
        type: "runPropertyChange",
        info: { id: id + 2, author: AUTHOR, date: DATE },
        previousFormatting: { italic: true },
        currentFormatting: { bold: true },
      },
    ],
  });
  const content = [schema.text(`before-${index}`)];
  switch (item.inline) {
    case "plain":
      break;
    case "insertion":
      content.push(schema.text(" new", [inserted]));
      break;
    case "deletion":
      content.push(schema.text(" old", [deleted]));
      break;
    case "mixed":
      content.push(schema.text(" old", [deleted]), schema.text(" new", [inserted]));
      break;
    case "property":
      content.push(schema.text(" format", [schema.marks.bold.create(), formatting]));
      break;
    default:
      item.inline satisfies never;
  }
  const attrs = {
    paraId: `${(index + 1).toString(16).padStart(8, "0")}`,
    ...(item.mark === "none" ? {} : { pPrMark: paragraphMark(item.mark, id + 3) }),
    ...(item.paragraphProperty
      ? {
          alignment: "right",
          _originalFormatting: { alignment: "right" },
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: id + 4, author: AUTHOR, date: DATE },
              previousFormatting: { alignment: "left" },
            },
          ],
        }
      : {}),
    ...(item.sectionProperty
      ? {
          _sectionProperties: {
            sectionStart: "continuous",
            propertyChanges: [
              {
                type: "sectionPropertyChange",
                info: { id: id + 5, author: AUTHOR, date: DATE },
                previousProperties: { sectionStart: "nextPage" },
              },
            ],
          },
        }
      : {}),
  };
  return schema.node("paragraph", attrs, content);
};

const table = (rowMarker: "none" | "trIns" | "trDel", cellMarker: "none" | "ins" | "del") => {
  const cell = (text: string, marker: "none" | "ins" | "del") =>
    schema.node(
      "tableCell",
      marker === "none"
        ? null
        : { cellMarker: { kind: marker, info: revision(text === "changed" ? 904 : 905) } },
      [
        schema.node(
          "paragraph",
          { paraId: text === "changed" ? "00000388" : "00000389" },
          schema.text(text),
        ),
      ],
    );
  const row = (text: string, marker: "none" | "trIns" | "trDel") =>
    schema.node("tableRow", marker === "none" ? null : { [marker]: revision(901) }, [
      cell(text, "none"),
      cell(`${text}-second`, "none"),
    ]);
  return schema.node("table", null, [
    row("stable", "none"),
    row("changed", rowMarker),
    schema.node("tableRow", null, [cell("changed", cellMarker), cell("stable", "none")]),
  ]);
};

const snapshot = (state: EditorState) => {
  const tracker = getChangeTrackerState(state);
  return {
    doc: state.doc.toJSON(),
    selection: state.selection.toJSON(),
    tracker: tracker && {
      ...tracker,
      changedParaIds: [...tracker.changedParaIds].toSorted(),
    },
    undoDepth: undoDepth(state),
  };
};

const run = (state: EditorState, command: Command) => {
  let transaction: Transaction | null = null;
  const result = command(state, (dispatched) => {
    expect(transaction).toBeNull();
    transaction = dispatched;
  });
  return {
    result,
    state: transaction ? state.apply(transaction) : state,
    dispatched: transaction !== null,
  };
};

const expectEquivalent = (state: EditorState) => {
  for (const mode of ["accept", "reject"] as const) {
    const bulk = run(state, mode === "accept" ? acceptAllChanges() : rejectAllChanges());
    const legacy = run(
      state,
      mode === "accept"
        ? acceptChange(0, state.doc.content.size)
        : rejectChange(0, state.doc.content.size),
    );
    expect(bulk.result).toBe(legacy.result);
    expect(bulk.dispatched).toBe(legacy.dispatched);
    expect(snapshot(bulk.state)).toEqual(snapshot(legacy.state));
    expect(() => bulk.state.doc.check()).not.toThrow();
  }
};

describe("resolve-all command equivalence", () => {
  test("matches the legacy whole-document range over mixed paragraph and table revisions", () => {
    fc.assert(
      fc.property(
        fc.array(paragraphCase, { minLength: 2, maxLength: 6 }),
        fc.constantFrom("none", "trIns", "trDel"),
        fc.constantFrom("none", "ins", "del"),
        (items, rowMarker, cellMarker) => {
          const blocks = items.map(paragraph);
          blocks.splice(1, 0, table(rowMarker, cellMarker));
          const doc = schema.node("doc", null, blocks);
          const state = EditorState.create({
            schema,
            doc,
            selection: TextSelection.create(doc, 2),
            plugins: [trackerPlugin],
          });
          expectEquivalent(state);
        },
      ),
      { seed: 2_609_260, numRuns: 32, verbose: true },
    );
  });

  test("matches paragraph joins, table-adjacent removal, and tracked section endpoints", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        { paraId: "000000a1", pPrMark: paragraphMark("del", 101) },
        schema.text("gone", [schema.marks.deletion.create(revision(102))]),
      ),
      schema.node("paragraph", { paraId: "000000a2" }, schema.text("following")),
      schema.node(
        "paragraph",
        {
          paraId: "000000a3",
          pPrMark: paragraphMark("del", 103),
          _sectionProperties: { sectionStart: "continuous" },
        },
        schema.text("removed", [schema.marks.deletion.create(revision(104))]),
      ),
      table("trIns", "del"),
      schema.node("paragraph", { paraId: "000000a4" }, schema.text("tail")),
    ]);
    expectEquivalent(EditorState.create({ schema, doc, plugins: [trackerPlugin, history()] }));
  });

  test("matches a pending vertical split and a row revision in the same table", () => {
    const cell = (text: string, attrs?: Record<string, unknown>) =>
      schema.node("tableCell", attrs, [schema.node("paragraph", null, schema.text(text))]);
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, schema.text("leading")),
      schema.node("table", null, [
        schema.node("tableRow", null, [cell("top")]),
        schema.node("tableRow", null, [
          cell("restored", {
            cellMarker: {
              kind: "merge",
              info: { revisionId: 201, author: AUTHOR, date: null },
              verticalMergeOriginal: "continue",
            },
          }),
        ]),
        schema.node("tableRow", { trIns: revision(202) }, [cell("new row")]),
      ]),
      schema.node("paragraph", null, schema.text("trailing")),
    ]);
    expectEquivalent(EditorState.create({ schema, doc, plugins: [trackerPlugin] }));
  });

  test("matches a pending vertical merge stored in a collapsed continuation cell", () => {
    const continuation = {
      type: "tableCell" as const,
      formatting: { vMerge: "continue" as const },
      structuralChange: {
        type: "tableCellMerge" as const,
        info: { id: 301, author: AUTHOR },
        verticalMerge: "continue" as const,
      },
      content: [
        {
          _docxParagraphSourceBinding: { type: "authored" as const },
          type: "paragraph" as const,
          content: [],
        },
      ],
    };
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, schema.text("before")),
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { rowspan: 2, _docxVMergeContinuationCells: [continuation] }, [
            schema.node("paragraph", null, schema.text("top")),
          ]),
        ]),
        schema.node("tableRow"),
        schema.node("tableRow", { trDel: revision(302) }, [
          schema.node("tableCell", null, [schema.node("paragraph", null, schema.text("old row"))]),
        ]),
      ]),
      schema.node("paragraph", null, schema.text("after")),
    ]);
    expectEquivalent(EditorState.create({ schema, doc, plugins: [trackerPlugin] }));
  });

  test("matches table, row, and cell property-only revisions", () => {
    const info = (id: number) => ({ id, author: AUTHOR, date: DATE });
    const changedCell = schema.node(
      "tableCell",
      {
        backgroundColor: "99CCFF",
        tcPrChange: [{ info: info(403), previousFormatting: { backgroundColor: "FFFFFF" } }],
      },
      [schema.node("paragraph", { paraId: "00000403" }, schema.text("cell"))],
    );
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: "00000400" }, schema.text("before")),
      schema.node(
        "table",
        {
          width: 7200,
          tblPrChange: [{ info: info(401), previousFormatting: { width: 6400 } }],
        },
        [
          schema.node(
            "tableRow",
            {
              height: 480,
              trPrChange: [{ info: info(402), previousFormatting: { height: 240 } }],
            },
            [
              changedCell,
              schema.node("tableCell", null, [
                schema.node("paragraph", { paraId: "00000404" }, schema.text("peer")),
              ]),
            ],
          ),
        ],
      ),
      schema.node("paragraph", { paraId: "00000405" }, schema.text("after")),
    ]);
    expectEquivalent(EditorState.create({ schema, doc, plugins: [trackerPlugin] }));
  });

  test("matches the no-revision dispatch contract", () => {
    const doc = schema.node("doc", null, [schema.node("paragraph", null, schema.text("plain"))]);
    expectEquivalent(EditorState.create({ schema, doc, plugins: [trackerPlugin] }));
  });
});
