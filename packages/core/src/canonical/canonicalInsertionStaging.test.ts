import { panic } from "better-result";
import { expect, spyOn, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import * as native from "@stll/docx-core/ops";
import {
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { canonicalOperationStagingAccess } from "../internal/canonicalOperationStaging";
import { readModel } from "../types/documentModel";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

test("insertion staging owns the verified operation and rejects foreign or stale results", () => {
  const session = createCanonicalSession({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content: [] }],
      },
    },
  }).unwrap();
  const source = readModel(session.document);
  const blockId =
    session.projection.firstParagraph?.blockId ??
    panic("The fixture has no paragraph anchor.");
  const state = EditorState.create({ doc: session.projection.doc });
  const selection = session.projection.selectionAt(state).unwrap();
  const text = { type: "text" as const, text: "Verified." };
  const block = {
    type: "paragraph",
    paraId: blockId,
    content: [{ type: "run", content: [text] }],
  } satisfies InsertBlocksOp["blocks"][number];
  const op = {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId },
    blocks: [block],
  } satisfies InsertBlocksOp;
  expect(() =>
    Reflect.apply(session.beginVerifiedOperationStaging, session, [Symbol()]),
  ).toThrow("Only the public-operation compiler");
  const staging = session.beginVerifiedOperationStaging(
    canonicalOperationStagingAccess,
  );
  expect(staging.apply(source, op).isErr()).toBe(true);
  block.paraId = "001ABCDE";
  const verified = staging.apply(source, op).unwrap();
  expect(staging.append({ ...verified })).toBe(false);
  block.paraId = "001ABCDF";
  text.text = "Mutated.";
  expect(staging.append(verified)).toBe(true);
  const commit = staging.prepare(state, selection).unwrap();
  expect(commit.ops).toEqual([
    {
      ...op,
      blocks: [
        {
          ...block,
          paraId: "001ABCDE",
          content: [
            { type: "run", content: [{ type: "text", text: "Verified." }] },
          ],
        },
      ],
    },
  ]);
  expect(commit.touched.inserted).toEqual(["001ABCDE"]);
  publishCanonicalProjection({ session, state, commit }).unwrap();
  const stale = staging.prepare(state, selection);
  if (stale.isOk()) panic("The previous version's staging should be stale.");
  expect(stale.error.message).toContain("stale canonical document");
});

test("one verified run contributes each owned effect to the journal exactly once", () => {
  const anchor = "00000001";
  const session = createCanonicalSession({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: anchor, content: [] }],
      },
    },
  }).unwrap();
  const document = readModel(session.document);
  const state = EditorState.create({ doc: session.projection.doc });
  const selection = session.projection.selectionAt(state).unwrap();
  const operations = Array.from({ length: 32 }, (_, index): InsertBlocksOp => {
    const id = 1 + 8 * index;
    return {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: anchor },
      blocks: [
        {
          type: "paragraph",
          paraId: native.formatParaId(index + 2),
          content: [
            { type: "run", content: [{ type: "text", text: String(index) }] },
          ],
        },
      ],
      revision: { id, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
      newIds: { revision: [id + 1, id + 2, id + 3] },
    };
  });
  const reference = native.applyDocumentOps(document, operations).unwrap();
  const staged = session.beginVerifiedOperationStaging(
    canonicalOperationStagingAccess,
  );
  const first = operations[0] ?? panic("Expected a first insertion");
  let attempts = 0;
  expect(
    staged.tryAppendInsertionGroup(document, () => {
      attempts += 1;
      return attempts === 1
        ? { type: "member", op: first }
        : { type: "fallback" };
    }),
  ).toEqual({ type: "fallback", reason: "compiler" });
  expect(attempts).toBe(2);
  const combined = spyOn(native, "combineEditEffects");
  try {
    const members = operations.values();
    const added = staged.tryAppendInsertionGroup(document, () => {
      const member = members.next();
      return member.done
        ? { type: "done" }
        : { type: "member", op: member.value };
    });
    if (added.type !== "materialized")
      panic("Expected a verified insertion run");
    expect(added.revisions.flat()).toEqual(reference.revisions);
    expect(combined).toHaveBeenCalledTimes(0);
    expect(
      staged.tryAppendInsertionGroup(document, () =>
        panic("Stale compiler ran"),
      ),
    ).toEqual({
      type: "fallback",
      reason: "base",
    });
    const commit = staged.prepare(state, selection).unwrap();
    expect(combined).toHaveBeenCalledTimes(1);
    expect(combined.mock.calls[0]?.[1].length).toBe(operations.length);
    expect(commit.document).toEqual(reference.document);
    expect(commit.touched).toEqual(reference.touched);
    expect(commit.ops).toEqual(operations);
    const published = publishCanonicalProjection({
      session,
      state,
      commit,
    }).unwrap();
    const undo = session.prepareUndo(published.state).unwrap();
    expect(undo.ops).toEqual(reference.inverse);
    expect(undo.document).toEqual(document);
  } finally {
    combined.mockRestore();
  }
});
