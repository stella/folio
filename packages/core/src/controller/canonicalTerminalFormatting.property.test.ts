import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  paragraphVisibleText,
  type DocumentOp,
} from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(30_000));

test("canonical tracked terminal deletion preserves the preceding formatting request", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(fc.constantFrom("Launch 🚀", "Tail", "東京", "é"), { minLength: 2, maxLength: 5 }),
      fc.constantFrom("start", "end", "center"),
      async (texts, alignment) => {
        const document = createEmptyDocument({ initialText: "" });
        document.package.document.content = texts.map((text, index) => ({
          type: "paragraph",
          paraId: (index + 1).toString(16).padStart(8, "0"),
          content: [{ type: "run", content: [{ type: "text", text }] }],
        }));
        const targetId = (texts.length - 1).toString(16).padStart(8, "0");
        const terminalId = texts.length.toString(16).padStart(8, "0");
        const revision = { id: 10, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
        const ops = [
          {
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: OP_STORIES.MAIN,
            blockId: targetId,
            patch: { alignment },
            revision,
          },
          {
            type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
            story: OP_STORIES.MAIN,
            blockIds: [terminalId],
            revision: { ...revision, id: 20 },
            newIds: { revision: [21, 22, 23, 24, 25] },
          },
        ] satisfies DocumentOp[];
        const session = createCanonicalSession(document).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const initial = session.document;
        state = publishCanonicalProjection({
          state,
          session,
          commit: session.prepareOperations(state, ops).unwrap(),
        }).unwrap().state;
        const suggested = session.document;
        state = publishCanonicalProjection({
          state,
          session,
          commit: session
            .prepareResolve(state, {
              revisionIds: [10, 20, 21, 22, 23, 24, 25],
              resolution: "accept",
            })
            .unwrap(),
        }).unwrap().state;
        const accepted = session.document.package.document.content;
        expect(
          accepted.flatMap((block) =>
            block.type === "paragraph" ? [paragraphVisibleText(block)] : [],
          ),
        ).toEqual(texts.slice(0, -1));
        const survivor = accepted.at(-1);
        expect(survivor?.type === "paragraph" && survivor.formatting?.alignment).toBe(alignment);
        state = publishCanonicalProjection({
          state,
          session,
          commit: session.prepareUndo(state).unwrap(),
        }).unwrap().state;
        expect(session.document).toEqual(suggested);
        state = publishCanonicalProjection({
          state,
          session,
          commit: session
            .prepareResolve(state, {
              revisionIds: [10, 20, 21, 22, 23, 24, 25],
              resolution: "reject",
            })
            .unwrap(),
        }).unwrap().state;
        expect(session.document).toEqual(initial);
        expect(state.doc.eq(session.projection.doc)).toBe(true);
      },
    ),
    { numRuns: 24 },
  );
});
