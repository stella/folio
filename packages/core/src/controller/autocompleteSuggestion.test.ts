/**
 * The inline suggestion owner: its lifecycle, how it follows edits (typing
 * its next characters, a peer's edit, any other edit), word-wise accepting,
 * and where a suggestion may start. The canonical cases commit real edits
 * to a session and feed the owner the session's committed batches, as the
 * hidden editor does.
 */

import { panic } from "better-result";
import { describe, expect, spyOn, test } from "bun:test";
import { Schema } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import * as documentOps from "@stll/docx-core/ops";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type TextPosition,
} from "@stll/docx-core/ops";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { readModel } from "../types/documentModel";
import {
  canTriggerAutocomplete,
  classifyAutocompleteEdit,
  createAutocompleteSuggestion,
  IDLE_AUTOCOMPLETE,
  legacyAutocompleteEdit,
  nextAutocompleteWord,
} from "./autocompleteSuggestion";

const schema = new Schema({
  nodes: {
    doc: { content: "block*" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { paraId: { default: null } },
      toDOM: () => ["p", 0],
    },
    heading: { group: "block", content: "inline*", toDOM: () => ["h1", 0] },
    code_block: {
      group: "block",
      content: "text*",
      marks: "",
      code: true,
      toDOM: () => ["pre", ["code", 0]],
    },
    text: { group: "inline" },
  },
});

type BlockName = "paragraph" | "heading" | "code_block";

const blockDoc = (block: BlockName, text: string): PMNode =>
  schema.node("doc", null, [
    schema.node(block, null, text.length > 0 ? [schema.text(text)] : []),
  ]);

const emptyDoc = (): PMNode => schema.node("doc", null, []);

// Caret immediately after the text of a single-block document: the
// block opens at doc position 0, so content starts at 1.
const endOfText = (text: string): number => 1 + text.length;

const mkState = (doc: PMNode, caret?: number): EditorState =>
  caret === undefined
    ? EditorState.create({ doc })
    : EditorState.create({ doc, selection: TextSelection.create(doc, caret) });

const at = (offset: number, blockId = "P1"): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

describe("autocomplete suggestion lifecycle", () => {
  test("streams tokens of its own request and ignores others", () => {
    const owner = createAutocompleteSuggestion();
    const seen: string[] = [];
    owner.subscribe((state) => seen.push(state.status));
    owner.start(at(3), "r1");
    expect(owner.append("r1", "lorem")).toBe(true);
    expect(owner.append("stale", "x")).toBe(false);
    expect(owner.append("r1", " ipsum")).toBe(true);
    expect(owner.get()).toEqual({
      status: "streaming",
      requestId: "r1",
      anchor: at(3),
      text: "lorem ipsum",
    });
    expect(owner.finish("r1")).toBe(true);
    expect(owner.get().status).toBe("shown");
    owner.clear();
    expect(owner.get()).toBe(IDLE_AUTOCOMPLETE);
    expect(seen).toEqual([
      "streaming",
      "streaming",
      "streaming",
      "shown",
      "idle",
    ]);
  });

  test("a suggestion that ends without text is dismissed", () => {
    const owner = createAutocompleteSuggestion();
    owner.start(at(0), "r1");
    expect(owner.finish("r1")).toBe(true);
    expect(owner.get()).toBe(IDLE_AUTOCOMPLETE);
    expect(owner.append("r1", "late")).toBe(false);
  });
});

describe("autocomplete suggestion edits", () => {
  const shown = (text = "next words here") => {
    const owner = createAutocompleteSuggestion();
    owner.start(at(5), "r1");
    owner.append("r1", text);
    owner.finish("r1");
    return owner;
  };

  test("typing its next characters at the anchor consumes them", () => {
    const owner = shown();
    owner.noteEdit({ type: "typed", at: at(5), text: "next " });
    expect(owner.get()).toMatchObject({ anchor: at(10), text: "words here" });
    owner.noteEdit({ type: "typed", at: at(10), text: "words here" });
    expect(owner.get()).toBe(IDLE_AUTOCOMPLETE);
  });

  test("typing anything else, or elsewhere, dismisses it", () => {
    const other = shown();
    other.noteEdit({ type: "typed", at: at(5), text: "x" });
    expect(other.get()).toBe(IDLE_AUTOCOMPLETE);
    const elsewhere = shown();
    elsewhere.noteEdit({ type: "typed", at: at(2), text: "n" });
    expect(elsewhere.get()).toBe(IDLE_AUTOCOMPLETE);
    const structural = shown();
    structural.noteEdit({ type: "other" });
    expect(structural.get()).toBe(IDLE_AUTOCOMPLETE);
  });

  /** A peer's edits around a shown suggestion, and where its anchor goes. */
  const expectPeerEditsMoveTheAnchor = () => {
    const insert = (offset: number, text: string): DocumentOp => ({
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at(offset),
      text,
      runProps: INHERIT_RUN_PROPS,
    });
    const owner = shown();
    owner.noteEdit({ type: "moved", ops: [insert(0, "abc")] });
    expect(owner.get()).toMatchObject({ anchor: at(8), status: "shown" });
    owner.noteEdit({ type: "moved", ops: [insert(8, "zz")] });
    expect(owner.get()).toMatchObject({ anchor: at(8), status: "shown" });
  };

  test("a peer's text before the anchor moves it; at the anchor lands after it", () => {
    expectPeerEditsMoveTheAnchor();
  });

  test("the peer check fails when the anchor is not moved over the edit", () => {
    // Turn the mapping off: every position comes back where it was.
    const transform = spyOn(
      documentOps,
      "transformPosition",
    ).mockImplementation((position) => position);
    try {
      expect(expectPeerEditsMoveTheAnchor).toThrow();
    } finally {
      transform.mockRestore();
    }
  });

  test("the next word takes its leading space and one space after it", () => {
    expect(nextAutocompleteWord(" next words")).toBe(" next ");
    expect(nextAutocompleteWord("last")).toBe("last");
  });

  test("a default-session insertion at one point reads as typing", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: "P1" }, [schema.text("Hello")]),
    ]);
    const state = EditorState.create({ doc });
    expect(legacyAutocompleteEdit(state.tr.insertText("!", 6))).toEqual({
      type: "typed",
      at: at(5),
      text: "!",
    });
    expect(legacyAutocompleteEdit(state.tr.delete(1, 2))).toEqual({
      type: "other",
    });
  });
});

/** A document from Markdown whose paragraphs carry ids, as a parsed DOCX's do. */
const withParagraphIds = (markdown: string) => {
  const document = fromMarkdown(markdown);
  for (const [index, block] of document.package.document.content.entries())
    if (block.type === "paragraph")
      block.paraId = (index + 1).toString(16).toUpperCase().padStart(8, "0");
  return document;
};

describe("autocomplete suggestion in a canonical session", () => {
  const openCanonical = () => {
    const session = createCanonicalSession(
      withParagraphIds("Dear reader, hello.\n\nSecond paragraph."),
    ).unwrap();
    let state = EditorState.create({ doc: session.projection.doc });
    const owner = createAutocompleteSuggestion();
    session.onCommitted((batch) =>
      owner.noteEdit(
        classifyAutocompleteEdit({ ops: batch.ops, origin: batch.origin }),
      ),
    );
    const [firstId] = readModel(
      session.document,
    ).package.document.content.flatMap((block) =>
      block.type === "paragraph" && block.paraId !== undefined
        ? [block.paraId]
        : [],
    );
    const first = (offset: number) => at(offset, firstId ?? panic("No id"));
    const typeAt = (offset: number, text: string) => {
      const position = session.projection.positionAt(first(offset)).unwrap();
      const commit = session
        .prepareReplace(state, { from: position, to: position, text })
        .unwrap();
      state = state.apply(commit.transaction);
      expect(commit.publish().isOk()).toBe(true);
    };
    let revision = 0;
    const peerTypes = (offset: number, text: string) => {
      revision += 1;
      const commit = session
        .prepareRemote(state, {
          type: "rebased",
          batch: {
            schema: DOCUMENT_OP_SCHEMA_VERSION,
            opId: `peer-${String(revision)}`,
            actor: "peer",
            baseRev: revision - 1,
            revision,
            ops: [
              {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: first(offset),
                text,
                runProps: INHERIT_RUN_PROPS,
              },
            ],
          },
        })
        .unwrap();
      state = state.apply(commit.transaction);
      expect(commit.publish().isOk()).toBe(true);
    };
    return { session, owner, first, typeAt, peerTypes };
  };

  test("typing the ghost's next characters moves it along; a peer's edit keeps it on its text", () => {
    const canonical = openCanonical();
    const { owner, first } = canonical;
    owner.start(first(19), "r1");
    owner.append("r1", " Welcome back.");
    owner.finish("r1");
    canonical.typeAt(19, " W");
    expect(owner.get()).toMatchObject({
      anchor: first(21),
      text: "elcome back.",
    });
    canonical.peerTypes(0, "Oh, ");
    expect(owner.get()).toMatchObject({
      anchor: first(25),
      text: "elcome back.",
    });
    // The suggestion is view state: it adds nothing to the history.
    const version = canonical.session.version;
    owner.append("r1", "");
    expect(canonical.session.version).toBe(version);
    canonical.typeAt(3, "x");
    expect(owner.get()).toBe(IDLE_AUTOCOMPLETE);
  });
});

describe("canTriggerAutocomplete", () => {
  test("fires after a word boundary in prose", () => {
    const state = mkState(blockDoc("paragraph", "Hello "), endOfText("Hello "));
    expect(canTriggerAutocomplete(state)).toEqual({ ok: true });
  });

  test("skips when the selection is a range", () => {
    const doc = blockDoc("paragraph", "hello");
    const state = EditorState.create({
      doc,
      selection: TextSelection.create(doc, 1, 4),
    });
    expect(canTriggerAutocomplete(state)).toEqual({
      ok: false,
      reason: "selection-non-empty",
    });
  });

  test("skips an empty document", () => {
    const state = mkState(emptyDoc());
    expect(canTriggerAutocomplete(state)).toEqual({
      ok: false,
      reason: "empty-doc",
    });
  });

  test("skips inside a heading dead zone", () => {
    const state = mkState(blockDoc("heading", "Title"), 1);
    expect(canTriggerAutocomplete(state)).toEqual({
      ok: false,
      reason: "deadzone",
      detail: "heading",
    });
  });

  test("skips inside a code block dead zone", () => {
    const state = mkState(blockDoc("code_block", "x = 1"), 1);
    expect(canTriggerAutocomplete(state)).toEqual({
      ok: false,
      reason: "deadzone",
      detail: "code_block",
    });
  });

  test("skips mid-word after an ASCII letter", () => {
    const state = mkState(blockDoc("paragraph", "foo"), endOfText("foo"));
    expect(canTriggerAutocomplete(state)).toEqual({
      ok: false,
      reason: "midword",
      detail: "o",
    });
  });

  test("treats an accented letter as a word continuation char", () => {
    const state = mkState(blockDoc("paragraph", "Dobrý"), endOfText("Dobrý"));
    expect(canTriggerAutocomplete(state)).toEqual({
      ok: false,
      reason: "midword",
      detail: "ý",
    });
  });
});
