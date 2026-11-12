import { expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  type IdLease,
  idLeaseHolds,
  introducedIds,
  leasedIdCensus,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import { readModel, type ReadonlyDocument } from "../types/documentModel";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";

const PARAGRAPH = "12345678";

const seed = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: PARAGRAPH,
          content: [
            { type: "run", content: [{ type: "text", text: "plain" }] },
          ],
        },
      ],
    },
  },
});

const lease: IdLease = {
  blocks: [
    { kind: "paragraph", start: 0x2000, end: 0x2010 },
    { kind: "revision", start: 7000, end: 7010 },
    { kind: "control", start: 7010, end: 7020 },
    { kind: "comment", start: 7020, end: 7030 },
  ],
};

const stateFor = (session: CanonicalSession, head = 3) => {
  const state = EditorState.create({ schema, doc: session.projection.doc });
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, head)),
  );
};

const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return next;
};

const minted = (before: ReadonlyDocument, after: ReadonlyDocument) =>
  introducedIds(
    leasedIdCensus(readModel(before)),
    leasedIdCensus(readModel(after)),
  );

test("a leased session mints new paragraph and revision ids from its lease, never twice", () => {
  const session = createCanonicalSession(seed(), undefined, {
    idLease: lease,
  }).unwrap();
  let state = stateFor(session);
  const original = session.document;
  state = accept(state, session.prepareSplit(state).unwrap());
  const split = minted(original, session.document);
  expect(split).toEqual([{ kind: "paragraph", id: 0x2000 }]);
  // The lease moved past the minted id, and undoing the split does not move it back.
  state = accept(state, session.prepareUndo(state).unwrap());
  expect(
    session.idLease?.blocks.find(({ kind }) => kind === "paragraph"),
  ).toEqual({ kind: "paragraph", start: 0x2001, end: 0x2010 });
  state = accept(state, session.prepareSplit(state).unwrap());
  expect(minted(original, session.document)).toEqual([
    { kind: "paragraph", id: 0x2001 },
  ]);
  session.setMode({ type: "suggesting", author: "Reviewer" });
  const before = session.document;
  state = accept(
    state,
    session
      .prepareReplace(state, {
        from: state.selection.head,
        to: state.selection.head,
        text: "X",
      })
      .unwrap(),
  );
  const suggested = minted(before, session.document);
  expect(suggested.length).toBeGreaterThan(0);
  for (const { kind, id } of suggested)
    expect(idLeaseHolds(lease, kind, id)).toBe(true);
});

test("without a lease a session mints the lowest free ids, as before", () => {
  const session = createCanonicalSession(seed()).unwrap();
  const state = stateFor(session);
  const original = session.document;
  accept(state, session.prepareSplit(state).unwrap());
  expect(minted(original, session.document)).toEqual([
    { kind: "paragraph", id: 1 },
  ]);
  expect(session.idLease).toBeUndefined();
});

test("a leased session refuses an edit introducing an id outside its lease", () => {
  const session = createCanonicalSession(seed(), undefined, {
    idLease: lease,
  }).unwrap();
  const state = stateFor(session);
  const refused = session.prepareOperations(state, [
    {
      type: "splitBlock",
      at: { story: OP_STORIES.MAIN, blockId: PARAGRAPH, offset: 2 },
      newBlockId: "00000001",
      newHalf: "second",
    },
  ]);
  expect(refused.isErr()).toBe(true);
  if (refused.isOk()) return;
  expect(refused.error.message).toContain("outside this author's id lease");
  // Granted blocks extend the lease; spent ids stay spent.
  expect(
    session
      .extendIdLease({ blocks: [{ kind: "paragraph", start: 1, end: 4 }] })
      .isOk(),
  ).toBe(true);
  expect(
    session
      .prepareOperations(state, [
        {
          type: "splitBlock",
          at: { story: OP_STORIES.MAIN, blockId: PARAGRAPH, offset: 2 },
          newBlockId: "00000001",
          newHalf: "second",
        },
      ])
      .isOk(),
  ).toBe(true);
});
