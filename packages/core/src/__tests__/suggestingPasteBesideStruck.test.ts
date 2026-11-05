/**
 * A paste beside text a tracked deletion struck, on either side of a paragraph
 * break: accepting reads as pasting into the text left, rejecting as no paste.
 */

import { describe, expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { shapeArrayBuffer } from "./documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
  textblocks,
} from "./editorHarness";

const texts = (state: EditorState) =>
  textblocks(state.doc).map(({ node }) => node.textContent);

const opened = async () =>
  createHarnessState(
    await parseShapeDocument(
      new Uint8Array(await shapeArrayBuffer("single-decimal-list")),
    ),
    "suggesting",
  );

const paragraphs = (state: EditorState, ...lines: string[]) =>
  Fragment.fromArray(
    lines.map((line) =>
      state.schema.nodes["paragraph"]!.create(null, state.schema.text(line)),
    ),
  );

/** "Alpha" whole and the "B" of "Beta" struck, with the caret after the "B". */
const struckAcrossBreak = async () => {
  const state = await opened();
  const blocks = textblocks(state.doc);
  const alpha = blocks.find(({ node }) => node.textContent === "Alpha");
  const beta = blocks.find(({ node }) => node.textContent === "Beta");
  if (!alpha || !beta) throw new Error("Fixture lost its list items");
  const view = new HeadlessEditorView(
    state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, alpha.pos + 1, beta.pos + 2),
      ),
    ),
  );
  view.pressKey("Delete");
  return { state, view, alpha, beta };
};

describe("a paste beside struck text", () => {
  test("two open paragraphs before a struck break keep their own break", async () => {
    const { state, view, alpha } = await struckAcrossBreak();
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, alpha.pos + 1),
      ),
    );
    view.paste(new Slice(paragraphs(state, "One", "Two"), 1, 1));

    const accepted = texts(resolveAllChanges(view.state, "accept"));
    expect(accepted).toContain("One");
    expect(accepted).toContain("Twoeta");
    expect(texts(resolveAllChanges(view.state, "reject"))).toEqual(
      texts(state),
    );
  });

  test("closed paragraphs after a struck prefix leave no empty paragraph", async () => {
    const { state, view } = await struckAcrossBreak();
    view.paste(new Slice(paragraphs(state, "One", "Two"), 0, 0));

    const accepted = texts(resolveAllChanges(view.state, "accept"));
    expect(accepted).toContain("One");
    expect(accepted).toContain("Two");
    expect(accepted).toContain("eta");
    expect(accepted).not.toContain("");
    expect(texts(resolveAllChanges(view.state, "reject"))).toEqual(
      texts(state),
    );
  });

  test("a second paste over selected cells leaves no paragraph the first one split off", async () => {
    const state = createHarnessState(
      await parseShapeDocument(
        new Uint8Array(await shapeArrayBuffer("tables")),
      ),
      "suggesting",
    );
    const cells: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type.name === "tableCell") cells.push(pos);
      return true;
    });
    // The last cell of a row with text of its own, alone.
    const target = cells.find(
      (pos) => state.doc.nodeAt(pos)?.textContent === "Cell D3",
    );
    if (target === undefined) throw new Error("Fixture lost its cell");
    const view = new HeadlessEditorView(
      state.apply(
        state.tr.setSelection(CellSelection.create(state.doc, target)),
      ),
    );
    const cellParagraphs = (current: EditorState) => {
      const lines: string[] = [];
      current.doc.nodeAt(target)?.forEach((child) => {
        lines.push(child.textContent);
      });
      return lines;
    };
    view.paste(new Slice(paragraphs(state, "One", "Two"), 1, 1));
    view.dispatch(
      view.state.tr.setSelection(CellSelection.create(view.state.doc, target)),
    );
    view.paste(new Slice(paragraphs(state, "Three"), 1, 1));

    expect(cellParagraphs(resolveAllChanges(view.state, "accept"))).toEqual([
      "Three",
    ]);
    expect(cellParagraphs(resolveAllChanges(view.state, "reject"))).toEqual([
      "Cell D3",
    ]);
  });

  test("a table over words across a break leaves the caret after them", async () => {
    const state = await opened();
    const blocks = textblocks(state.doc);
    const alpha = blocks.find(({ node }) => node.textContent === "Alpha");
    const beta = blocks.find(({ node }) => node.textContent === "Beta");
    if (!alpha || !beta) throw new Error("Fixture lost its list items");
    const { table, tableRow, tableCell, paragraph } = state.schema.nodes;
    const view = new HeadlessEditorView(
      state.apply(
        state.tr.setSelection(
          TextSelection.create(state.doc, alpha.pos + 1, beta.pos + 2),
        ),
      ),
    );
    view.paste(
      new Slice(
        Fragment.from(
          table!.create(
            null,
            tableRow!.create(
              null,
              tableCell!.create(
                null,
                paragraph!.create(null, state.schema.text("Cell")),
              ),
            ),
          ),
        ),
        0,
        0,
      ),
    );

    const { $head } = view.state.selection;
    expect($head.parent.textContent).toBe("Beta");
    expect($head.parentOffset).toBe(1);
  });
});
