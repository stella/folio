import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { schema } from "../prosemirror/schema";
import type { Document, TableCell } from "../types/document";

const cell = (paraId: string, text: string): TableCell => ({
  type: "tableCell",
  content: [
    {
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

const tableDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [cell("21000000", "A"), cell("22000000", "B")],
            },
          ],
        },
        {
          type: "paragraph",
          paraId: "30000000",
          content: [
            { type: "run", content: [{ type: "text", text: "After" }] },
          ],
        },
      ],
    },
  },
});

const texts = (document: Document) => {
  const table = document.package.document.content[0];
  if (table?.type !== "table") return panic("No table.");
  return table.rows.flatMap((row) =>
    row.cells.map((each) =>
      each.content
        .flatMap((block) =>
          block.type === "paragraph"
            ? block.content.flatMap((item) =>
                item.type === "run"
                  ? item.content.flatMap((leaf) =>
                      leaf.type === "text" ? [leaf.text] : [],
                    )
                  : [],
              )
            : [],
        )
        .join(""),
    ),
  );
};

for (const mode of ["editing", "suggesting"] as const)
  test(`typing over selected cells in ${mode} mode clears them and writes into the first`, () => {
    const session = createCanonicalSession(tableDocument()).unwrap();
    if (mode === "suggesting")
      session.setMode({ type: "suggesting", author: "Reviewer" });
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const cells: number[] = [];
    state.doc.descendants((node, position) => {
      if (node.type.name === "tableCell") cells.push(position);
      return true;
    });
    const [anchor, head] = cells;
    if (anchor === undefined || head === undefined) return panic("No cells.");
    state = state.apply(
      state.tr.setSelection(CellSelection.create(state.doc, anchor, head)),
    );
    const commit = session
      .prepareReplace(state, {
        from: state.selection.from,
        to: state.selection.to,
        text: "Z",
      })
      .unwrap();
    expect(commit.publish().isOk()).toBe(true);
    if (mode === "editing") {
      expect(texts(session.document)).toEqual(["Z", ""]);
      return;
    }
    // A suggestion records the cleared text as deleted and the typed text as
    // inserted in the first cell.
    const table = JSON.stringify(session.document.package.document.content[0]);
    expect(table).toMatch(/"insertion".*"text":"Z"/u);
    expect(table.match(/"deletion"/gu)).toHaveLength(2);
  });

test("a tab over selected cells clears them and puts the tab in the first", () => {
  const session = createCanonicalSession(tableDocument()).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const cells: number[] = [];
  state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell") cells.push(position);
    return true;
  });
  const [anchor, head] = cells;
  if (anchor === undefined || head === undefined) return panic("No cells.");
  state = state.apply(
    state.tr.setSelection(CellSelection.create(state.doc, anchor, head)),
  );
  const commit = session
    .prepareCellSelectionStructure(state, (at) => ({
      type: "insertAtom",
      from: at,
      to: at,
      atom: { type: "tab" },
    }))
    .unwrap();
  expect(commit.publish().isOk()).toBe(true);
  expect(texts(session.document)).toEqual(["", ""]);
  const table = JSON.stringify(session.document.package.document.content[0]);
  expect(table.match(/"type":"tab"/gu)).toHaveLength(1);
});

test("a paragraph break over selected cells clears them and splits the first", () => {
  const session = createCanonicalSession(tableDocument()).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const cells: number[] = [];
  state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell") cells.push(position);
    return true;
  });
  const [anchor, head] = cells;
  if (anchor === undefined || head === undefined) return panic("No cells.");
  state = state.apply(
    state.tr.setSelection(CellSelection.create(state.doc, anchor, head)),
  );
  const commit = session
    .prepareCellSelectionStructure(state, (at) => ({
      type: "replaceFragment",
      from: at,
      to: at,
      paragraphs: [
        { type: "paragraph", content: [] },
        { type: "paragraph", content: [] },
      ],
      openStart: 1,
      openEnd: 1,
    }))
    .unwrap();
  expect(commit.publish().isOk()).toBe(true);
  const table = session.document.package.document.content[0];
  if (table?.type !== "table") return panic("No table.");
  expect(table.rows[0]?.cells.map((each) => each.content.length)).toEqual([
    2, 1,
  ]);
  expect(texts(session.document)).toEqual(["", ""]);
});
