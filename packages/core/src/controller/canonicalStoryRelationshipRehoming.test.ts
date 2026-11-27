/**
 * A picture and a link read in the body and placed into a header name the
 * body's own picture and link there once saved and reopened, though the
 * header part's own relationships hold the same ids for its own: inserted by
 * an intent that says where its content was read, and pasted from the
 * body's own copy. Undo takes the edit and the relationships it gave the
 * header back exactly.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { OP_STORIES, type OpStory } from "@stll/docx-core/ops";

import { assertExactModel } from "../../../../test/exactModel";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import { readModel } from "../types/documentModel";
import {
  linkTarget,
  paragraphsHolding,
  pictureAndLink,
  sourceText,
  storyRelationshipPackage,
} from "./__fixtures__/storyRelationshipPackage";
import { prepareCanonicalPaste } from "./canonicalClipboard";

const HEADER: OpStory = { kind: "header", rId: "rIdHeader1" };

const read = (buffer: ArrayBuffer) =>
  parseDocx(buffer.slice(0), { preloadFonts: false, detectVariables: false });

const only = (paragraphs: readonly Paragraph[]): Paragraph =>
  paragraphs.length === 1
    ? (paragraphs[0] ?? panic("A paragraph is missing."))
    : panic(`Expected one paragraph, found ${paragraphs.length}.`);

/** The range of the first projected paragraph whose text holds `text`. */
const projectedParagraph = (state: EditorState, text: string) => {
  let range: { from: number; to: number } | undefined;
  state.doc.descendants((node, position) => {
    if (range !== undefined || node.type.name !== "paragraph") return true;
    if (node.textContent.includes(text))
      range = { from: position + 1, to: position + node.nodeSize - 1 };
    return false;
  });
  return range ?? panic(`No projected paragraph holds "${text}".`);
};

type Placement = "insert" | "paste";

/** Place the body's source paragraph into the header; the commit and its state. */
const place = (
  document: Document,
  session: CanonicalSession,
  placement: Placement,
): { commit: CanonicalCommit; state: EditorState } => {
  const header = EditorState.create({
    schema,
    doc: session.projectStory(HEADER).unwrap().doc,
  });
  if (placement === "insert") {
    const source = only(
      paragraphsHolding(document, OP_STORIES.MAIN, sourceText("body")),
    );
    const anchor = only(paragraphsHolding(document, HEADER, "Header text"));
    return {
      state: header,
      commit: session
        .prepareIntents(header, {
          intents: [
            {
              type: "insertParagraphs",
              story: HEADER,
              at: {
                type: "after",
                blockId:
                  anchor.paraId ?? panic("The header paragraph has no id."),
              },
              paragraphs: [structuredClone(source)],
            },
          ],
          contentStory: OP_STORIES.MAIN,
        })
        .unwrap(),
    };
  }
  const body = EditorState.create({ schema, doc: session.projection.doc });
  const copied = projectedParagraph(body, sourceText("body"));
  const slice = TextSelection.create(
    body.doc,
    copied.from,
    copied.to,
  ).content();
  const state = header.apply(
    header.tr.setSelection(
      TextSelection.create(
        header.doc,
        projectedParagraph(header, "Header text").to,
      ),
    ),
  );
  return {
    state,
    commit: prepareCanonicalPaste({
      session,
      state,
      slice,
      sourceDocument: readModel(session.document),
      sourceStory: OP_STORIES.MAIN,
      story: HEADER,
    }).unwrap(),
  };
};

describe("content placed from the body into a header", () => {
  for (const placement of ["insert", "paste"] as const)
    test(`a body picture and link the header takes by ${placement} resolve to the body's own after a save and reopen, and undo restores the document`, async () => {
      const opened = await read(await storyRelationshipPackage());
      const session = createCanonicalSession(opened).unwrap();
      const before = session.captureSaveSnapshot().document;
      const bodySource = pictureAndLink(
        only(paragraphsHolding(opened, OP_STORIES.MAIN, sourceText("body"))),
      );
      const headerSource = pictureAndLink(
        only(paragraphsHolding(opened, HEADER, sourceText("header"))),
      );
      expect(bodySource.image.src).not.toBe(headerSource.image.src);
      const { commit, state } = place(opened, session, placement);
      const after = state.apply(commit.transaction);
      commit.publish().unwrap();
      for (const selectiveSave of [false, true]) {
        // oxlint-disable-next-line no-await-in-loop -- each save is read back in turn
        const saved = await serializeCanonicalSave({
          snapshot: session.captureSaveSnapshot(),
          featureFlags: { selectiveSave },
        });
        // oxlint-disable-next-line no-await-in-loop -- each save is read back in turn
        const reopened = await read(saved.buffer);
        const placed = pictureAndLink(
          only(paragraphsHolding(reopened, HEADER, sourceText("body"))),
        );
        expect(placed.image.src).toBe(bodySource.image.src);
        expect(placed.href).toBe(linkTarget("body"));
        const own = pictureAndLink(
          only(paragraphsHolding(reopened, HEADER, sourceText("header"))),
        );
        expect(own.image.src).toBe(headerSource.image.src);
        expect(own.href).toBe(linkTarget("header"));
      }
      session.prepareUndo(after, HEADER).unwrap().publish().unwrap();
      assertExactModel(session.captureSaveSnapshot().document, before);
    });
});
