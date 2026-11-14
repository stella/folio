/**
 * Where invisible markers (comment ranges, bookmarks, permission ranges,
 * proofing marks) share one editor position, that position stands for the
 * source gap typed text takes there (`defaultInsertionGap`). A caret the
 * editor places and a caret an edit returns then name the same gap, so
 * characters typed one after another stay together.
 */
import { describe, expect, test } from "bun:test";
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
import { parseDocx } from "../docx/parser";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";

const SHAPES = ["comments", "invisible-markers"] as const;

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

/** The text a comment covers once every change is accepted, and the paragraph's text. */
const replaceAndAccept = (
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
    test(`${mode}: replacing exactly the commented words keeps the comment on the new text`, () => {
      const result = replaceAndAccept(mode, (text) => {
        const from = text.indexOf("commented");
        return { from, to: from + "commented".length };
      });
      expect(result).toEqual({ covered: "xy", text: "Lead xy tail" });
    });

    test(`${mode}: replacing a paragraph's whole text takes its markers, so the comment is emptied`, () => {
      const result = replaceAndAccept(mode, (text) => ({
        from: 0,
        to: text.length,
      }));
      expect(result).toEqual({ covered: "", text: "xy" });
    });
  }
});
