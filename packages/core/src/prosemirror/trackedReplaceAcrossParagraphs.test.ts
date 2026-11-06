/**
 * A key binding or command that replaces a selection across paragraphs builds
 * an ordinary transaction. In suggesting mode the selected text stays, struck,
 * and what replaced it follows: rejecting gives the document back, accepting
 * gives what the same edit gives outside suggesting mode.
 */

import { describe, expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";
import { AllSelection, TextSelection } from "prosemirror-state";

import { documentShape } from "../__tests__/documentShapes";
import {
  createHarnessState,
  type EditorMode,
  harnessManager,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  resolveAllChanges,
  summarizeState,
} from "../__tests__/editorHarness";
import {
  buildPlainTextSlice,
  pastePlainTextSlice,
} from "./commands/pastePlainText";
import { dispatchEditorTextInput } from "./textInput";

type Edit = { id: string; run: (view: HeadlessEditorView) => unknown };

const EDITS: readonly Edit[] = [
  { id: "Tab", run: (view) => view.pressKey("Tab") },
  { id: "Shift-Enter", run: (view) => view.pressKey("Shift-Enter") },
  {
    id: "insertHyperlink",
    run: (view) =>
      view.execute(
        harnessManager().requireCommand("insertHyperlink")(
          "a link",
          "https://example.org/",
        ),
      ),
  },
  {
    id: "plain text paste",
    run: (view) =>
      pastePlainTextSlice(
        view as never,
        buildPlainTextSlice("one\ntwo", view.state.schema),
      ),
  },
  {
    id: "delete range",
    run: (view) => {
      const { from, to } = view.state.selection;
      view.dispatch(view.state.tr.deleteRange(from, to));
    },
  },
];

const SHAPES = ["plain-markdown", "comments", "notes", "image"] as const;
const PLACEMENTS = ["cross-paragraph", "document"] as const;

const open = async (
  shapeId: string,
  placement: (typeof PLACEMENTS)[number],
  mode: EditorMode,
) => {
  const shape = documentShape(shapeId);
  const base = await parseShapeDocument(await shape.build());
  const state = createHarnessState(base, mode);
  const before =
    placement === "document"
      ? state.apply(state.tr.setSelection(new AllSelection(state.doc)))
      : placeSelection(state, shape.focus, placement);
  if (!before) throw new Error("No selection");
  return before;
};

const edited = (before: EditorState, edit: Edit): EditorState => {
  const view = new HeadlessEditorView(before);
  edit.run(view);
  expect(view.state.doc.eq(before.doc)).toBe(false);
  return view.state;
};

const CASES = EDITS.flatMap((edit) =>
  SHAPES.flatMap((shape) =>
    PLACEMENTS.map((placement) => [edit.id, shape, placement, edit] as const),
  ),
);

describe("replacing a selection across paragraphs in suggesting mode", () => {
  test.each(CASES)(
    "%s over %s, %s: rejecting restores, accepting matches editing",
    async (_id, shape, placement, edit) => {
      const before = await open(shape, placement, "suggesting");
      const suggesting = edited(before, edit);
      const direct = edited(await open(shape, placement, "editing"), edit);
      expect(summarizeState(resolveAllChanges(suggesting, "reject"))).toEqual(
        summarizeState(resolveAllChanges(before, "reject")),
      );
      expect(summarizeState(resolveAllChanges(suggesting, "accept"))).toEqual(
        summarizeState(resolveAllChanges(direct, "accept")),
      );
    },
  );

  test("a replacement from a paragraph into a table restores on rejecting", async () => {
    const shape = documentShape("tables");
    const base = await parseShapeDocument(await shape.build());
    const state = createHarnessState(base, "suggesting");
    const blocks: number[] = [];
    state.doc.descendants((node, position) => {
      if (node.isTextblock && node.content.size > 0) blocks.push(position + 1);
      return true;
    });
    const [first = 0, second = 0] = blocks;
    const before = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, first + 1, second + 2),
      ),
    );
    const after = edited(before, EDITS[0] as Edit);
    expect(summarizeState(resolveAllChanges(after, "reject"))).toEqual(
      summarizeState(resolveAllChanges(before, "reject")),
    );
  });

  test("a list marker typed over the whole document makes a list, as in editing", async () => {
    const typeMarker = (before: EditorState) => {
      const view = new HeadlessEditorView(before);
      for (const character of "- ") {
        dispatchEditorTextInput(view as never, character);
      }
      return view.state;
    };
    const paragraphs = (state: EditorState) => {
      const found: { text: string; listed: boolean }[] = [];
      state.doc.descendants((node) => {
        if (node.isTextblock) {
          found.push({
            text: node.textContent,
            listed: Boolean(node.attrs["numPr"]),
          });
        }
        return true;
      });
      return found;
    };
    const before = await open("plain-markdown", "document", "suggesting");
    const suggesting = typeMarker(before);
    expect(paragraphs(resolveAllChanges(suggesting, "accept"))).toEqual(
      paragraphs(
        typeMarker(await open("plain-markdown", "document", "editing")),
      ),
    );
    expect(summarizeState(resolveAllChanges(suggesting, "reject"))).toEqual(
      summarizeState(resolveAllChanges(before, "reject")),
    );
  });

  test("a list marker after the text a joined paragraph keeps stays text", async () => {
    const before = await open(
      "plain-markdown",
      "cross-paragraph",
      "suggesting",
    );
    const view = new HeadlessEditorView(before);
    for (const character of "- ") {
      dispatchEditorTextInput(view as never, character);
    }
    const joined = resolveAllChanges(view.state, "accept").doc.children.find(
      (paragraph) => paragraph.textContent.includes("- "),
    );
    expect(joined).toBeDefined();
    expect(joined?.attrs["numPr"] ?? null).toBeNull();
  });

  test("the caret follows what replaced the selection", async () => {
    const before = await open(
      "plain-markdown",
      "cross-paragraph",
      "suggesting",
    );
    const after = edited(before, EDITS[0] as Edit);
    const { $from, empty } = after.selection;
    expect(empty).toBe(true);
    expect(after.selection).toBeInstanceOf(TextSelection);
    expect($from.nodeBefore?.type.name).toBe("tab");
  });
});
