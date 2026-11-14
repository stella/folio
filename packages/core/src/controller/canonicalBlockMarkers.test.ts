import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { BlockContent, Document, Paragraph } from "../types/document";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

type Marker = Extract<
  BlockContent,
  {
    type:
      | "bookmarkStart"
      | "bookmarkEnd"
      | "moveFromRangeStart"
      | "moveFromRangeEnd";
  }
>;

/** A bookmark and a move range that open and close between paragraphs. */
const markedDocument = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("10000000", "Before"),
        { type: "bookmarkStart", id: 3, name: "Clause" } as Marker,
        paragraph("20000000", "Inside"),
        { type: "bookmarkEnd", id: 3 } as Marker,
        paragraph("30000000", "After"),
      ],
    },
  },
});

const session = async () =>
  createCanonicalSession(
    await parseDocx(
      (
        await prepareCanonicalDocxInput(await createDocx(markedDocument()))
      ).unwrap(),
      { preloadFonts: false, detectVariables: false },
    ),
  ).unwrap();

const stateFor = (current: CanonicalSession) =>
  EditorState.create({ schema, doc: current.projection.doc });
const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return next;
};
const at = (state: EditorState, text: string, offset = text.length) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.isText && node.text === text) found = position + offset;
    return true;
  });
  return found ?? panic(`No text ${text}`);
};
const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

/** Block kinds and texts in order, as a reader of the saved file sees them. */
const outline = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) =>
            item.type === "run"
              ? item.content.flatMap((leaf) =>
                  leaf.type === "text" ? [leaf.text] : [],
                )
              : [item.type],
          )
          .join("")
      : block.type,
  );

const reopened = async (current: CanonicalSession) => {
  const saved = await serializeCanonicalSave({
    snapshot: current.captureSaveSnapshot(),
    featureFlags: { selectiveSave: true },
  });
  return parseDocx(saved.buffer, { preloadFonts: false });
};

describe("bookmark markers between paragraphs in the canonical session", () => {
  test("open, and typing beside them keeps each marker once, in place, through save", async () => {
    const current = await session();
    let state = stateFor(current);
    const end = at(state, "Inside");
    state = select(state, end);
    state = accept(
      state,
      current.prepareReplace(state, { from: end, to: end, text: "!" }).unwrap(),
    );
    expect(outline(await reopened(current))).toEqual([
      "Before",
      "bookmarkStart",
      "Inside!",
      "bookmarkEnd",
      "After",
    ]);
  });

  test("Enter beside a marker keeps it where it was, and undo is exact", async () => {
    const current = await session();
    const before = current.document;
    let state = stateFor(current);
    state = select(state, at(state, "Inside", 3));
    state = accept(state, current.prepareSplit(state).unwrap());
    expect(outline(current.document)).toEqual([
      "Before",
      "bookmarkStart",
      "Ins",
      "ide",
      "bookmarkEnd",
      "After",
    ]);
    state = accept(state, current.prepareUndo(state).unwrap());
    expect(current.document).toEqual(before);
  });

  test("a join across a marker refuses and leaves the document as it was", async () => {
    const current = await session();
    const before = current.document;
    let state = stateFor(current);
    state = select(state, at(state, "After", 0));
    // The engine joins only paragraphs adjacent in their block list.
    const joined = current.prepareJoin(state, "backward");
    expect(joined.isErr()).toBe(true);
    expect(current.document).toBe(before);
  });
});
