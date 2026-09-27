/**
 * Ordinary editing across a section break.
 *
 * Deleting the mark of the paragraph that ends a section deletes the break:
 * the section's content joins the following section and takes its
 * properties, as deleting the paragraph as a block does. The change tracker
 * carries each break to the paragraph whose mark survived and records the
 * removal the save is to accept; any other loss of a break stays unrecorded.
 */

import { describe, expect, test } from "bun:test";
import { deleteSelection, joinBackward, joinForward } from "prosemirror-commands";
import { Schema, Slice } from "prosemirror-model";
import { EditorState, TextSelection, type Transaction } from "prosemirror-state";

import type { SectionProperties } from "../types/document";
import {
  getTrackedSectionEndpointRemoval,
  ParagraphChangeTrackerExtension,
} from "./extensions/features/ParagraphChangeTrackerExtension";
import { sectionPropertiesOf } from "./sectionCarrier";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      content: "inline*",
      group: "block",
      attrs: { _sectionProperties: { default: null } },
    },
    text: { group: "inline" },
  },
});

const trackerPlugin = ParagraphChangeTrackerExtension().onSchemaReady({ schema }).plugins?.at(0);
if (!trackerPlugin) throw new Error("the tracker has no plugin");

const sectionOne: SectionProperties = {
  sectionStart: "nextPage",
  headerReferences: [{ type: "first", rId: "rIdOne" }],
};
const sectionTwo: SectionProperties = {
  sectionStart: "nextPage",
  orientation: "landscape",
  headerReferences: [{ type: "even", rId: "rIdTwo" }],
};

const paragraph = (text: string, record: SectionProperties | null = null) =>
  schema.node("paragraph", { _sectionProperties: record }, text ? schema.text(text) : []);

/** "One." ends section one, "Two." ends section two, "Three." runs to the end. */
const threeSections = () =>
  EditorState.create({
    schema,
    doc: schema.node("doc", null, [
      paragraph("One.", sectionOne),
      paragraph("Two.", sectionTwo),
      paragraph("Three."),
    ]),
    plugins: [trackerPlugin],
  });

/** Where the text of the paragraph at `index` starts. */
const textStart = (state: EditorState, index: number): number => {
  let position = 1;
  for (let at = 0; at < index; at++) position += state.doc.child(at).nodeSize;
  return position;
};

const run = (
  state: EditorState,
  selection: { anchor: number; head?: number },
  command: (state: EditorState, dispatch: (tr: Transaction) => void) => boolean,
): EditorState => {
  let next = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, selection.anchor, selection.head)),
  );
  const before = next;
  expect(
    command(before, (tr) => {
      next = next.apply(tr);
    }),
  ).toBe(true);
  return next;
};

const paragraphs = (state: EditorState) => {
  const out: { text: string; record: SectionProperties | null }[] = [];
  state.doc.forEach((node) =>
    out.push({ text: node.textContent, record: sectionPropertiesOf(node) }),
  );
  return out;
};

describe("editing across a section break", () => {
  test("Backspace at the start of the next paragraph merges the section into the next", () => {
    const state = threeSections();
    const next = run(state, { anchor: textStart(state, 1) }, joinBackward);

    expect(paragraphs(next)).toEqual([
      { text: "One.Two.", record: sectionTwo },
      { text: "Three.", record: null },
    ]);
    expect(getTrackedSectionEndpointRemoval(next)).toMatchObject({
      sourceParagraphEndpointCount: 2,
      expectedParagraphEndpointCount: 1,
      removedReferences: [{ part: "header", type: "first", relationshipId: "rIdOne" }],
    });
  });

  test("Delete at the end of a section-ending paragraph does the same", () => {
    const state = threeSections();
    const next = run(state, { anchor: textStart(state, 1) - 2 }, joinForward);

    expect(paragraphs(next)[0]).toEqual({ text: "One.Two.", record: sectionTwo });
    expect(getTrackedSectionEndpointRemoval(next)?.expectedParagraphEndpointCount).toBe(1);
  });

  test("a join into a section-ending paragraph keeps its break", () => {
    const state = threeSections();
    const next = run(state, { anchor: textStart(state, 2) }, joinBackward);

    expect(paragraphs(next)).toEqual([
      { text: "One.", record: sectionOne },
      { text: "Two.Three.", record: null },
    ]);
    // Section two's mark was deleted; its content joins the last section.
    expect(getTrackedSectionEndpointRemoval(next)?.removedReferences).toEqual([
      { part: "header", type: "even", relationshipId: "rIdTwo" },
    ]);
  });

  test("a selection deleted across two breaks removes both", () => {
    const state = threeSections();
    const next = run(
      state,
      { anchor: textStart(state, 0) + 2, head: textStart(state, 2) + 2 },
      deleteSelection,
    );

    expect(paragraphs(next)).toEqual([{ text: "Onree.", record: null }]);
    expect(getTrackedSectionEndpointRemoval(next)).toMatchObject({
      sourceParagraphEndpointCount: 2,
      expectedParagraphEndpointCount: 0,
    });
  });

  test("a paste over a break leaves the pasted paragraphs without it", () => {
    const state = threeSections();
    const slice = new Slice(
      schema.node("doc", null, [paragraph("A"), paragraph("B")]).content,
      1,
      1,
    );
    const next = run(
      state,
      { anchor: textStart(state, 0) + 2, head: textStart(state, 1) + 2 },
      (current, dispatch) => {
        dispatch(current.tr.replaceSelection(slice));
        return true;
      },
    );

    expect(paragraphs(next)).toEqual([
      { text: "OnA", record: null },
      { text: "Bo.", record: sectionTwo },
      { text: "Three.", record: null },
    ]);
    expect(getTrackedSectionEndpointRemoval(next)?.expectedParagraphEndpointCount).toBe(1);
  });

  test("an undo that restores the break takes its removal back", () => {
    const state = threeSections();
    const joined = run(state, { anchor: textStart(state, 1) }, joinBackward);
    const restored = joined.apply(
      joined.tr.replaceWith(0, joined.doc.content.size, state.doc.content),
    );

    expect(paragraphs(restored).map(({ record }) => record)).toEqual([
      sectionOne,
      sectionTwo,
      null,
    ]);
    expect(getTrackedSectionEndpointRemoval(restored)).toBeNull();
  });

  test("clearing a break without deleting its mark stays unrecorded", () => {
    const state = threeSections();
    const next = state.apply(state.tr.setNodeAttribute(0, "_sectionProperties", null));

    expect(getTrackedSectionEndpointRemoval(next)).toBeNull();
  });
});
