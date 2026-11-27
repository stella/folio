import { describe, expect, test } from "bun:test";
import {
  AllSelection,
  EditorState,
  type Selection,
  TextSelection,
} from "prosemirror-state";
import type { DecorationSet } from "prosemirror-view";

import { schema } from "../schema";
import { createHiddenEditorLayoutWindowPlugin } from "./hiddenEditorLayoutWindow";

const BLOCKS = 12;

const doc = schema.node(
  "doc",
  null,
  Array.from({ length: BLOCKS }, (_, index) =>
    schema.node("paragraph", null, [schema.text(`Block ${String(index)}`)]),
  ),
);

/** Where the text of block `index` starts. */
const inBlock = (index: number): number => {
  let position = 0;
  for (let at = 0; at < index; at += 1) position += doc.child(at).nodeSize;
  return position + 1;
};

const plugin = createHiddenEditorLayoutWindowPlugin();
const stateAt = (selection: (state: EditorState) => Selection) => {
  const state = EditorState.create({ doc, plugins: [plugin] });
  return state.apply(state.tr.setSelection(selection(state)));
};

const decorationsOf = (state: EditorState): DecorationSet => {
  const decorations = plugin.props.decorations?.call(plugin, state);
  if (
    decorations === undefined ||
    decorations === null ||
    !("find" in decorations)
  )
    throw new Error("the plugin gives no decoration set");
  return decorations as DecorationSet;
};

/** The indexes of the blocks the plugin marks laid out. */
const laidOut = (state: EditorState): number[] =>
  decorationsOf(state)
    .find()
    .map((decoration) => state.doc.resolve(decoration.from).index(0))
    .sort((a, b) => a - b);

describe("hidden editor layout window", () => {
  test("a caret lays out its block and the block on either side", () => {
    expect(
      laidOut(
        stateAt((state) => TextSelection.create(state.doc, inBlock(5) + 2)),
      ),
    ).toEqual([4, 5, 6]);
  });

  test("a range lays out every block it reaches, and one either side", () => {
    expect(
      laidOut(
        stateAt((state) =>
          TextSelection.create(state.doc, inBlock(3) + 1, inBlock(7) + 1),
        ),
      ),
    ).toEqual([2, 3, 4, 5, 6, 7, 8]);
  });

  test("the window stops at the first and last blocks", () => {
    expect(
      laidOut(stateAt((state) => TextSelection.create(state.doc, inBlock(0)))),
    ).toEqual([0, 1]);
    expect(
      laidOut(
        stateAt((state) =>
          TextSelection.create(state.doc, inBlock(BLOCKS - 1)),
        ),
      ),
    ).toEqual([BLOCKS - 2, BLOCKS - 1]);
  });

  test("selecting everything lays out every block", () => {
    expect(laidOut(stateAt((state) => new AllSelection(state.doc)))).toEqual(
      Array.from({ length: BLOCKS }, (_, index) => index),
    );
  });

  test("a caret moving inside its block keeps the same window", () => {
    const first = stateAt((state) =>
      TextSelection.create(state.doc, inBlock(5) + 1),
    );
    const moved = first.apply(
      first.tr.setSelection(TextSelection.create(first.doc, inBlock(5) + 4)),
    );
    expect(decorationsOf(moved)).toBe(decorationsOf(first));
  });

  test("an edit moves the window with the text", () => {
    const first = stateAt((state) =>
      TextSelection.create(state.doc, inBlock(5) + 1),
    );
    const edited = first.apply(first.tr.insertText("typed ", inBlock(1)));
    expect(laidOut(edited)).toEqual([4, 5, 6]);
  });
});
