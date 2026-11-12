/**
 * A list pasted while suggesting arrives as closed paragraphs. The paragraph
 * the paste lands in keeps its original break on the last pasted paragraph
 * (every break inside the pasted range is marked inserted), so that paragraph
 * records the properties the break had: rejecting the paste then restores the
 * paragraph as it was, not as a list item.
 */
import { describe, expect, test } from "bun:test";
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { DOCUMENT_SHAPES } from "./documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
  summarizeReviewer,
} from "./editorHarness";

const pastedList = (doc: PMNode) => {
  const { schema } = doc.type;
  const item = (text: string, level: number) =>
    schema.node(
      "paragraph",
      { _pastedHtmlList: { group: 1, kind: "bullet", level } },
      schema.text(text),
    );
  return new Slice(
    Fragment.from([item("Alpha", 0), item("Nested", 1), item("Omega", 0)]),
    1,
    1,
  );
};

const pastedParagraphs = (doc: PMNode) => {
  const { schema } = doc.type;
  return new Slice(
    Fragment.from([
      schema.node("paragraph", null, schema.text("Title")),
      schema.node("paragraph", null, schema.text("Body")),
    ]),
    1,
    1,
  );
};

/** Text and kind of each block, as a reader of the package sees it. */
const blocksOf = async (bytes: ArrayBuffer) =>
  summarizeReviewer(await FolioDocxReviewer.fromBuffer(bytes)).map(
    ({ kind, text }) => ({ kind, text }),
  );

const SELECTIONS = {
  "a caret inside the paragraph": (start: number) => [start + 10, start + 10],
  "a caret at the end of the paragraph": (start: number, end: number) => [
    end,
    end,
  ],
  "words running to the end of the paragraph": (start: number, end: number) => [
    start + 25,
    end,
  ],
} as const satisfies Record<string, (start: number, end: number) => number[]>;

describe("a list pasted while suggesting, then rejected", () => {
  for (const [name, select] of Object.entries(SELECTIONS))
    for (const followUp of [false, true])
      test(`over ${name}${followUp ? ", then Backspace and another paste" : ""}: restores the paragraphs`, async () => {
        const shape = DOCUMENT_SHAPES.find(({ id }) => id === "notes");
        if (!shape) throw new Error("no notes shape");
        const bytes = await shape.build();
        const base = await parseShapeDocument(bytes);
        const view = new HeadlessEditorView(
          createHarnessState(base, "suggesting"),
        );
        let start = -1;
        let end = -1;
        view.state.doc.descendants((node, pos) => {
          if (node.isTextblock && node.textContent.startsWith("A sentence")) {
            start = pos + 1;
            end = pos + node.nodeSize - 1;
          }
          return start < 0;
        });
        const [from = start, to = from] = select(start, end);
        view.dispatch(
          view.state.tr.setSelection(
            TextSelection.create(view.state.doc, from, to),
          ),
        );
        view.paste(pastedList(view.state.doc));
        if (followUp) {
          view.pressKey("Backspace");
          view.paste(pastedParagraphs(view.state.doc));
        }
        const saved = await saveHarnessState(view.state, base);
        const rejecting = await FolioDocxReviewer.fromBuffer(
          saved.bytes.buffer as ArrayBuffer,
        );
        rejecting.rejectAll();
        expect(await blocksOf(await rejecting.toBuffer())).toEqual(
          await blocksOf(bytes.buffer as ArrayBuffer),
        );
      });
});
