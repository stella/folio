/**
 * Where invisible markers (comment ranges, bookmarks, permission ranges,
 * proofing marks) share one editor position, that position stands for the
 * source gap typed text takes there (`defaultInsertionGap`). A caret the
 * editor places and a caret an edit returns then name the same gap, so
 * characters typed one after another stay together.
 */
import { afterEach, describe, expect, setSystemTime, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  defaultInsertionGap,
  OP_STORIES,
  paragraphLogicalText,
} from "@stll/docx-core/ops";

import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";

const SHAPES = ["comments", "invisible-markers"] as const;
afterEach(() => setSystemTime());

for (const id of SHAPES)
  test(`${id}: every insertion gap has an editor position, and it reads back as itself`, async () => {
    const shape = DOCUMENT_SHAPES.find((entry) => entry.id === id);
    if (shape === undefined) throw new TypeError(`Missing shape ${id}`);
    const document = await parseDocx(await shape.build(), {
      preloadFonts: false,
    });
    const session = createCanonicalSession(document).unwrap();
    const projection = session.projectStory(OP_STORIES.MAIN).unwrap();
    for (const block of document.package.document.content) {
      if (block.type !== "paragraph" || block.paraId === undefined) continue;
      const length = paragraphLogicalText(block).length;
      for (let offset = 0; offset <= length; offset += 1) {
        const address = {
          story: OP_STORIES.MAIN,
          blockId: block.paraId,
          offset,
          zeroWidthBefore: defaultInsertionGap(block.content, offset)
            .zeroWidthBefore,
        };
        const read = projection
          .addressAt(projection.positionAt(address).unwrap())
          .unwrap();
        expect(read.offset).toBe(offset);
        expect(read.zeroWidthBefore ?? 0).toBe(address.zeroWidthBefore);
      }
    }
  });

const replacementState = (
  mode: "editing" | "suggesting",
  select: (text: string) => { from: number; to: number },
) => {
  const document: Document = {
    package: {
      document: {
        comments: [
          {
            id: 0,
            anchorKind: "range",
            author: "Reviewer",
            content: [
              {
                type: "paragraph",
                paraId: "20000000",
                content: [
                  { type: "run", content: [{ type: "text", text: "Note" }] },
                ],
              },
            ],
          },
        ],
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            content: [
              { type: "run", content: [{ type: "text", text: "Lead " }] },
              { type: "commentRangeStart", id: 0 },
              { type: "run", content: [{ type: "text", text: "commented" }] },
              { type: "commentRangeEnd", id: 0 },
              { type: "run", content: [{ type: "text", text: " tail" }] },
            ],
          },
        ],
      },
    },
  };
  const session = createCanonicalSession(document).unwrap();
  session.setMode(
    mode === "editing" ? { type: mode } : { type: mode, author: "R" },
  );
  let state = EditorState.create({ doc: session.projection.doc });
  const { from, to } = select(state.doc.textContent);
  state = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, 1 + from, 1 + to)),
  );
  return { session, state };
};

/** The text a comment covers once every change is accepted, and the paragraph's text. */
const replaceAndAccept = (
  mode: "editing" | "suggesting",
  select: (text: string) => { from: number; to: number },
) => {
  const prepared = replacementState(mode, select);
  const { session } = prepared;
  let { state } = prepared;
  for (const text of ["x", "y"]) {
    const commit = session
      .prepareReplace(state, {
        from: state.selection.from,
        to: state.selection.to,
        text,
        semantic: "typing",
      })
      .unwrap();
    state = publishCanonicalProjection({ session, state, commit }).unwrap()
      .state;
  }
  const { ids } = getTrackedChangeStatsFromDoc(toProseDoc(session.document));
  if (ids.length > 0) {
    const commit = session
      .prepareResolve(state, { revisionIds: ids, resolution: "accept" })
      .unwrap();
    state = publishCanonicalProjection({ session, state, commit }).unwrap()
      .state;
  }
  const projected = toProseDoc(session.document);
  let covered = "";
  projected.descendants((node) => {
    if (node.isText && node.marks.some((mark) => mark.type.name === "comment"))
      covered += node.text ?? "";
  });
  return { covered, text: projected.textContent };
};

describe("which markers a replaced selection takes", () => {
  for (const mode of ["editing", "suggesting"] as const) {
    test(`${mode}: replacing exactly the commented words leaves the new text outside the comment`, () => {
      const result = replaceAndAccept(mode, (text) => {
        const from = text.indexOf("commented");
        return { from, to: from + "commented".length };
      });
      // The comment's words all went, so it goes with them.
      expect(result).toEqual({ covered: "", text: "Lead xy tail" });
    });

    test(`${mode}: replacing a paragraph containing an internal comment refuses atomically`, async () => {
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const { session, state } = replacementState(mode, (text) => ({
        from: 0,
        to: text.length,
      }));
      const document = session.document;
      const snapshot = session.captureSaveSnapshot();
      const saved = await serializeCanonicalSave({ snapshot });
      const version = session.version;
      const history = { undo: session.canUndo, redo: session.canRedo };
      let commits = 0;
      session.onCommitted(() => commits++);
      const result = session.prepareReplace(state, {
        from: state.selection.from,
        to: state.selection.to,
        text: "x",
        semantic: "typing",
      });
      // Empty covered text hid a disagreement: direct editing kept a collapsed
      // definition, while accepting the suggestion removed it entirely.
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.ground).toEqual({
          gap: CANONICAL_GAP.commentReplacementAlignment,
        });
      expect(session.document).toBe(document);
      expect(session.version).toBe(version);
      expect({ undo: session.canUndo, redo: session.canRedo }).toEqual(history);
      expect(commits).toBe(0);
      expect(session.captureSaveSnapshot()).toEqual(snapshot);
      expect(
        new Uint8Array(
          (
            await serializeCanonicalSave({
              snapshot: session.captureSaveSnapshot(),
            })
          ).buffer,
        ),
      ).toEqual(new Uint8Array(saved.buffer));
    });
  }
});
