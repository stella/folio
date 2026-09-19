/**
 * Editing a neighbour must not move the markup folio keeps opaquely.
 *
 * The node records no index, so its position is whatever ProseMirror's own
 * mapping says it is. These are the three edits that would break an index:
 * inserting a block in front of it, splitting the block in front of it, and
 * deleting the block in front of it. In each case the markup must still stand
 * between the same surviving neighbours — and when the block it followed is
 * gone, it attaches to the next survivor rather than disappearing with it.
 */

import { describe, expect, test } from "bun:test";
import { EditorState, type Plugin, TextSelection } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";

import { schema } from "../../schema";
import { PreservedBlockExtension } from "./PreservedBlockExtension";

const PRESERVED_XML = '<w:permStart w:id="77" w:edGrp="everyone"/>';

const integrityPlugin = (): Plugin => {
  const plugin = PreservedBlockExtension().onSchemaReady({ schema }).plugins?.at(0);
  if (!plugin) {
    throw new Error("PreservedBlockExtension must enforce node integrity");
  }
  return plugin;
};

const paragraph = (text: string): PMNode =>
  schema.node("paragraph", null, text === "" ? [] : [schema.text(text)]);

/** `one` · the capture · `two`. */
const createState = (): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      paragraph("one"),
      schema.node("preservedBlock", { xml: PRESERVED_XML }),
      paragraph("two"),
    ]),
    plugins: [integrityPlugin()],
  });

/** The block sequence, with a capture shown as its markup. */
const shapeOf = (state: EditorState): string[] => {
  const shape: string[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  state.doc.forEach((node) => {
    shape.push(node.type.name === "preservedBlock" ? node.attrs["xml"] : node.textContent);
  });
  return shape;
};

/** The position just inside the block at index `index`. */
const insideBlock = (state: EditorState, index: number): number => {
  let position = 0;
  let found = -1;
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  state.doc.forEach((node, offset, currentIndex) => {
    if (currentIndex === index) {
      found = offset + 1;
    }
    position = offset + node.nodeSize;
  });
  if (found === -1) {
    throw new Error(`no block at index ${index} (doc ends at ${position})`);
  }
  return found;
};

describe("a preserved block keeps its place while its neighbours change", () => {
  test("inserting a block in front of it does not move it", () => {
    const state = createState();
    const transaction = state.tr.insert(0, paragraph("zero"));
    expect(shapeOf(state.apply(transaction))).toEqual(["zero", "one", PRESERVED_XML, "two"]);
  });

  test("splitting the block in front of it leaves it after both halves", () => {
    const state = createState();
    // Between `o` and `ne` of the first paragraph.
    const transaction = state.tr.split(insideBlock(state, 0) + 1);
    expect(shapeOf(state.apply(transaction))).toEqual(["o", "ne", PRESERVED_XML, "two"]);
  });

  test("deleting the block in front of it attaches it to the next survivor", () => {
    const state = createState();
    const first = state.doc.firstChild;
    if (!first) {
      throw new Error("the fixture must open with a paragraph");
    }
    const transaction = state.tr.delete(0, first.nodeSize);
    expect(shapeOf(state.apply(transaction))).toEqual([PRESERVED_XML, "two"]);
  });

  test("merging the blocks around it keeps it between them", () => {
    const state = createState();
    // Typing at the end of `two` is the ordinary edit a reviewer makes next to
    // a protected range; the range must not move because of it.
    const end = insideBlock(state, 2) + 3;
    const transaction = state.tr.insertText("!", end);
    expect(shapeOf(state.apply(transaction))).toEqual(["one", PRESERVED_XML, "two!"]);
  });

  test("the integrity plugin drops a node whose markup did not survive a paste", () => {
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        paragraph("one"),
        schema.node("preservedBlock", { xml: "" }),
        paragraph("two"),
      ]),
      plugins: [integrityPlugin()],
    });
    // The plugin runs on a changed document, which is the state a paste leaves.
    expect(shapeOf(state.apply(state.tr.insertText("!", insideBlock(state, 0) + 3)))).toEqual([
      "one!",
      "two",
    ]);
  });

  test("it takes no selection and shows nothing", () => {
    const spec = schema.nodes["preservedBlock"]?.spec;
    expect(spec?.selectable).toBe(false);
    expect(spec?.atom).toBe(true);

    const node = schema.node("preservedBlock", { xml: PRESERVED_XML });
    const toDOM = node.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("PreservedBlockExtension must define toDOM");
    }
    expect(toDOM(node)).toEqual([
      "div",
      {
        "data-docx-preserved-block": PRESERVED_XML,
        "aria-hidden": "true",
        contenteditable: "false",
        style: "display: none;",
      },
    ]);
  });

  test("a text selection never lands on it", () => {
    const state = createState();
    const selection = TextSelection.near(state.doc.resolve(insideBlock(state, 0)));
    expect(selection.$from.parent.type.name).toBe("paragraph");
  });
});
