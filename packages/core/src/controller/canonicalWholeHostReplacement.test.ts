import { afterEach, expect, setSystemTime, test } from "bun:test";
import { AllSelection, EditorState, TextSelection } from "prosemirror-state";

import {
  applyDocumentOps,
  type CreateCommentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  reviewFieldsOf,
} from "@stll/docx-core/ops";

import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import {
  type CanonicalCommittedBatch,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";

const DATE = "2026-01-02T03:04:05Z";
afterEach(() => setSystemTime());

const fixture = async () => {
  const shape = DOCUMENT_SHAPES.find(({ id }) => id === "invisible-markers");
  if (!shape) throw new TypeError("The actual marker fixture exists");
  const prepared = (
    await prepareCanonicalDocxInput(await shape.build())
  ).unwrap();
  const source = await parseDocx(prepared, {
    preloadFonts: false,
    detectVariables: false,
  });
  const host = source.package.document.content.at(0);
  if (host?.type !== "paragraph" || host.paraId === undefined)
    throw new TypeError("The actual source host is identified");
  const point = {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 20,
      author: "Reviewer",
      content: [
        {
          type: "paragraph",
          paraId: "70000020",
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Standalone point" }],
            },
          ],
        },
      ],
    },
    anchor: {
      kind: "point",
      at: { story: OP_STORIES.MAIN, blockId: host.paraId, offset: 1 },
    },
  } satisfies CreateCommentOp;
  const replies = [0, 20].map(
    (parentId) =>
      ({
        type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        comment: {
          id: 100 + parentId,
          author: "Reply",
          content: [
            {
              type: "paragraph",
              paraId: (0x70000100 + parentId).toString(16).toUpperCase(),
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: `Reply ${parentId}` }],
                },
              ],
            },
          ],
        },
        anchor: { kind: "reply", parentId },
      }) satisfies CreateCommentOp,
  );
  return applyDocumentOps(source, [point, ...replies]).unwrap().document;
};

for (const endpoint of ["whole", "partial-last"] as const)
  test(`${endpoint}: whole retired comment hosts give direct and accepted typing the same source-owned state`, async () => {
    setSystemTime(new Date(DATE));
    const document = await fixture();
    const outcomes = [];
    for (const mode of ["editing", "suggesting"] as const) {
      const session = createCanonicalSession(document).unwrap();
      session.setMode(
        mode === "editing"
          ? { type: mode }
          : { type: mode, author: "Reviewer" },
      );
      const original = cloneDocumentWithParagraphPropertySources(
        session.document,
      );
      const last = document.package.document.content.at(-1);
      if (last?.type !== "paragraph" || last.paraId === undefined)
        throw new TypeError("The final source paragraph is identified");
      let state = EditorState.create({ doc: session.projection.doc });
      const end =
        endpoint === "whole"
          ? state.doc.content.size
          : session.projection
              .positionAt({
                story: OP_STORIES.MAIN,
                blockId: last.paraId,
                offset: 2,
              })
              .unwrap();
      state = state.apply(
        state.tr.setSelection(
          endpoint === "whole"
            ? new AllSelection(state.doc)
            : TextSelection.create(state.doc, 1, end),
        ),
      );
      const batches: CanonicalCommittedBatch[] = [];
      const off = session.onCommitted((batch) => batches.push(batch));
      for (const text of ["x", "y", "z"]) {
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
      off();
      const typed = cloneDocumentWithParagraphPropertySources(session.document);
      expect(
        applyDocumentOps(
          typed,
          batches.toReversed().flatMap(({ inverse }) => inverse),
        ).unwrap().document,
      ).toStrictEqual(original);
      expect(
        applyDocumentOps(
          typed,
          JSON.parse(
            JSON.stringify(
              batches.toReversed().flatMap(({ inverse }) => inverse),
            ),
          ),
        ).unwrap().document,
      ).toStrictEqual(original);
      const ids = getTrackedChangeStatsFromDoc(state.doc).ids;
      if (mode === "suggesting") {
        const rejected = createCanonicalSession(typed).unwrap();
        let rejectedState = EditorState.create({
          doc: rejected.projection.doc,
        });
        const reject = rejected
          .prepareResolve(rejectedState, {
            revisionIds: ids,
            resolution: "reject",
          })
          .unwrap();
        rejectedState = publishCanonicalProjection({
          session: rejected,
          state: rejectedState,
          commit: reject,
        }).unwrap().state;
        expect(rejected.document).toStrictEqual(original);
        const accept = session
          .prepareResolve(state, { revisionIds: ids, resolution: "accept" })
          .unwrap();
        state = publishCanonicalProjection({
          session,
          state,
          commit: accept,
        }).unwrap().state;
      }
      expect(state.doc.textContent).toBe(
        endpoint === "whole" ? "xyz" : "xyzil.",
      );
      expect(session.document.package.document.content).toHaveLength(1);
      expect(session.document.package.document.content.at(0)?.type).toBe(
        "paragraph",
      );
      expect(session.projection.firstParagraph?.blockId).toBe(last.paraId);
      expect(
        session
          .getCommittedComments()
          .map(({ id }) => id)
          .sort((a, b) => a - b),
      ).toEqual([3, 20, 120]);
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
      });
      const reopened = createCanonicalSession(
        await parseDocx(saved.buffer, {
          preloadFonts: false,
          detectVariables: false,
        }),
      ).unwrap();
      expect(reopened.projection.doc.textContent).toBe(state.doc.textContent);
      expect(
        reopened
          .getCommittedComments()
          .map(({ id, parentId }) => ({ id, parentId })),
      ).toEqual(
        session
          .getCommittedComments()
          .map(({ id, parentId }) => ({ id, parentId })),
      );
      const final = cloneDocumentWithParagraphPropertySources(session.document);
      outcomes.push(final);
      const historyBound = session.version;
      for (let step = 0; session.canUndo; step++) {
        expect(step).toBeLessThan(historyBound);
        const undo = session.prepareUndo(state).unwrap();
        state = publishCanonicalProjection({
          session,
          state,
          commit: undo,
        }).unwrap().state;
      }
      expect(session.document).toStrictEqual(original);
      for (let step = 0; session.canRedo; step++) {
        expect(step).toBeLessThan(historyBound);
        const redo = session.prepareRedo(state).unwrap();
        state = publishCanonicalProjection({
          session,
          state,
          commit: redo,
        }).unwrap().state;
      }
      expect(session.document).toStrictEqual(final);
    }
    expect(outcomes).toHaveLength(2);
    expect(outcomes.at(0)).toStrictEqual(outcomes.at(1));
  });

for (const boundary of [
  "partial-first",
  "surviving-host-comment",
  "section-endpoint",
  "pending-host-mark",
] as const)
  // Another author's pending break goes with its host in a direct edit, but
  // a suggestion cannot retract it.
  for (const mode of boundary === "pending-host-mark"
    ? (["suggesting"] as const)
    : (["editing", "suggesting"] as const))
    test(`${mode}/${boundary}: ${mode === "editing" && boundary === "surviving-host-comment" ? "typing over every word of the surviving host's comment takes it" : "unproven comment-host ownership refuses with all state unchanged"}`, async () => {
      setSystemTime(new Date(DATE));
      let document = await fixture();
      const first = document.package.document.content.at(0);
      const last = document.package.document.content.at(-1);
      if (
        first?.type !== "paragraph" ||
        first.paraId === undefined ||
        last?.type !== "paragraph" ||
        last.paraId === undefined
      )
        throw new TypeError("The fixture has identified endpoint paragraphs");
      switch (boundary) {
        case "partial-first":
          break;
        case "surviving-host-comment":
          document = applyDocumentOps(document, [
            {
              type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
              comment: {
                id: 30,
                author: "Reviewer",
                content: [
                  { type: "paragraph", paraId: "70000030", content: [] },
                ],
              },
              anchor: {
                kind: "range",
                from: {
                  story: OP_STORIES.MAIN,
                  blockId: last.paraId,
                  offset: 1,
                },
                to: { story: OP_STORIES.MAIN, blockId: last.paraId, offset: 4 },
              },
            },
          ]).unwrap().document;
          break;
        case "section-endpoint":
          document = applyDocumentOps(document, [
            {
              type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
              endpoint: { type: "paragraph", blockId: first.paraId },
              properties: {
                type: "present",
                value: { pageWidth: 12240, pageHeight: 15840 },
              },
            },
          ]).unwrap().document;
          break;
        case "pending-host-mark":
          document = applyDocumentOps(document, [
            {
              type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
              story: OP_STORIES.MAIN,
              blockId: first.paraId,
              expected: reviewFieldsOf(first),
              review: {
                ...reviewFieldsOf(first),
                pPrMark: {
                  kind: "ins",
                  info: { id: 501, author: "Writer", date: DATE },
                },
              },
            },
          ]).unwrap().document;
          break;
        default: {
          const unreachable: never = boundary;
          throw new TypeError(`Unexpected boundary ${unreachable}`);
        }
      }
      const session = createCanonicalSession(document).unwrap();
      session.setMode(
        mode === "editing"
          ? { type: mode }
          : { type: mode, author: "Reviewer" },
      );
      let state = EditorState.create({ doc: session.projection.doc });
      state = state.apply(
        state.tr.setSelection(
          boundary === "partial-first"
            ? TextSelection.create(state.doc, 2, state.doc.content.size - 1)
            : new AllSelection(state.doc),
        ),
      );
      const before = session.document;
      const version = session.version;
      const blocks = [...session.reservedBlockIds];
      // The writer may assign missing thread paragraph ids in its owned snapshot copy.
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
      });
      const snapshot = session.captureSaveSnapshot();
      let events = 0;
      const off = session.onCommitted(() => events++);
      const result = session.prepareReplace(state, {
        from: state.selection.from,
        to: state.selection.to,
        text: "xyz",
        semantic: "typing",
      });
      if (mode === "editing" && boundary === "surviving-host-comment") {
        // The selection takes every word of the surviving host's comment
        // too, so the direct edit takes the comment with them, as it takes
        // the hosts' own ranged threads; point threads keep their places.
        state = publishCanonicalProjection({
          session,
          state,
          commit: result.unwrap(),
        }).unwrap().state;
        expect(state.doc.textContent).toBe("xyz");
        expect(
          session.document.package.document.comments?.map(({ id }) => id),
        ).toEqual([3, 20, 120]);
        const undo = session.prepareUndo(state).unwrap();
        state = publishCanonicalProjection({
          session,
          state,
          commit: undo,
        }).unwrap().state;
        expect(session.document).toStrictEqual(before);
        off();
        return;
      }
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.ground).toEqual({
          gap: CANONICAL_GAP.commentReplacementAlignment,
        });
      expect(session.document).toBe(before);
      expect(session.version).toBe(version);
      expect([...session.reservedBlockIds]).toEqual(blocks);
      expect({ undo: session.canUndo, redo: session.canRedo, events }).toEqual({
        undo: false,
        redo: false,
        events: 0,
      });
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
      off();
    });

test("editing: a host's pending words and break retire with it, and undo restores them", async () => {
  setSystemTime(new Date(DATE));
  const source = await fixture();
  const first = source.package.document.content.at(0);
  if (first?.type !== "paragraph" || first.paraId === undefined)
    throw new TypeError("The fixture has its first host");
  const info = { id: 501, author: "Drafter", date: DATE };
  const document = applyDocumentOps(source, [
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story: OP_STORIES.MAIN,
      blockId: first.paraId,
      expected: reviewFieldsOf(first),
      review: { ...reviewFieldsOf(first), pPrMark: { kind: "ins", info } },
    },
    {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: { story: OP_STORIES.MAIN, blockId: first.paraId, offset: 0 },
      slice: {
        openStart: 0,
        openEnd: 0,
        content: [{ type: "run", content: [{ type: "text", text: "New " }] }],
      },
      revision: { ...info, id: 502 },
    },
  ]).unwrap().document;
  const session = createCanonicalSession(document).unwrap();
  session.setMode({ type: "editing" });
  const original = cloneDocumentWithParagraphPropertySources(session.document);
  let state = EditorState.create({ doc: session.projection.doc });
  expect(getTrackedChangeStatsFromDoc(state.doc).ids).toHaveLength(2);
  state = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
  const commit = session
    .prepareReplace(state, {
      from: state.selection.from,
      to: state.selection.to,
      text: "xyz",
      semantic: "typing",
    })
    .unwrap();
  state = publishCanonicalProjection({ session, state, commit }).unwrap().state;
  expect(state.doc.textContent).toBe("xyz");
  expect(getTrackedChangeStatsFromDoc(state.doc).ids).toEqual([]);
  expect(
    session
      .getCommittedComments()
      .map(({ id }) => id)
      .sort((a, b) => a - b),
  ).toEqual([3, 20, 120]);
  const undo = session.prepareUndo(state).unwrap();
  publishCanonicalProjection({ session, state, commit: undo }).unwrap();
  expect(session.document).toStrictEqual(original);
});
