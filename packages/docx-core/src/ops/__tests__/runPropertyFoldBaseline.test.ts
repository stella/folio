/**
 * An editor action folds a tracked run patch into the run's pending property
 * change, keeping the properties the change started from. A patch back to
 * those properties leaves nothing to review, so the change goes.
 */

import { describe, expect, test } from "bun:test";

import type { Document, Run } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  PROPERTY_REVIEW_POLICIES,
  type RunPropsPatch,
} from "../types";

const BLOCK_ID = "00000001";

/** "ab", made bold from unformatted as change 1. */
const PENDING_BOLD: Run = {
  type: "run",
  formatting: { bold: true },
  propertyChanges: [
    {
      type: "runPropertyChange",
      info: { id: 1, author: "A", date: "2026-01-01T00:00:00Z" },
      previousFormatting: {},
    },
  ],
  content: [{ type: "text", text: "ab" }],
};

const documentOf = (run: Run): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: BLOCK_ID, content: [run] }],
    },
  },
});

type PatchOptions = { from: number; to: number; patch: RunPropsPatch };

const patched = ({ from, to, patch }: PatchOptions) => {
  const applied = applyDocumentOp(documentOf(PENDING_BOLD), {
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: { story: OP_STORIES.MAIN, blockId: BLOCK_ID, offset: from },
    to: { story: OP_STORIES.MAIN, blockId: BLOCK_ID, offset: to },
    patch,
    propertyReview: PROPERTY_REVIEW_POLICIES.APPEND,
    revision: { id: 2, author: "B", date: "2026-01-02T00:00:00Z" },
    newIds: { revision: [3] },
  });
  if (applied.isErr()) throw applied.error;
  const [paragraph] = applied.value.document.package.document.content;
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return { paragraph, applied: applied.value };
};

describe("a tracked run patch over a pending property change", () => {
  test("back to where the change started leaves nothing pending", () => {
    const { paragraph, applied } = patched({
      from: 0,
      to: 2,
      patch: { bold: null },
    });
    // The run states what rejecting the change gives.
    const rejected = applyDocumentOp(documentOf(PENDING_BOLD), {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [1],
      decision: "reject",
    });
    if (rejected.isErr()) throw rejected.error;
    expect(paragraph).toEqual(
      rejected.value.document.package.document.content.at(0),
    );
    expect(paragraph.content).toEqual([
      { type: "run", formatting: {}, content: [{ type: "text", text: "ab" }] },
    ]);
    expect(applied.revisions).toEqual([]);
  });

  test("back to the start on part of the run keeps the change on the rest", () => {
    const { paragraph } = patched({ from: 0, to: 1, patch: { bold: null } });
    expect(paragraph.content).toMatchObject([
      { type: "run", content: [{ type: "text", text: "a" }] },
      {
        type: "run",
        formatting: { bold: true },
        propertyChanges: [{ info: { author: "A" }, previousFormatting: {} }],
        content: [{ type: "text", text: "b" }],
      },
    ]);
    expect(paragraph.content.at(0)).not.toHaveProperty("propertyChanges");
  });

  test("elsewhere folds into the change, keeping where it started", () => {
    const { paragraph } = patched({ from: 0, to: 2, patch: { italic: true } });
    expect(paragraph.content).toMatchObject([
      {
        type: "run",
        formatting: { bold: true, italic: true },
        propertyChanges: [{ info: { author: "B" }, previousFormatting: {} }],
      },
    ]);
  });
});

describe("the id of a folded run property change", () => {
  test("a fold over the whole run keeps the pending change's id", () => {
    const { paragraph, applied } = patched({
      from: 0,
      to: 2,
      patch: { italic: true },
    });
    expect(paragraph.content).toMatchObject([
      { propertyChanges: [{ info: { id: 1, author: "B" } }] },
    ]);
    // The kept id is no new revision; the patch reports it as folded.
    expect(applied.revisions).toEqual([]);
    expect(applied.folded).toEqual([1]);
  });

  test("a piece the patch cuts off takes the stamp's id; the rest keeps the change", () => {
    const { paragraph, applied } = patched({
      from: 0,
      to: 1,
      patch: { italic: true },
    });
    expect(paragraph.content).toMatchObject([
      {
        content: [{ type: "text", text: "a" }],
        formatting: { bold: true, italic: true },
        propertyChanges: [
          { info: { id: 2, author: "B" }, previousFormatting: {} },
        ],
      },
      {
        content: [{ type: "text", text: "b" }],
        formatting: { bold: true },
        propertyChanges: [{ info: { id: 1, author: "A" } }],
      },
    ]);
    expect(applied.revisions).toEqual([2]);
    // Change 1 still records what it did on the rest: nothing folded in.
    expect(applied.folded).toEqual([]);
  });

  test("a change the patch settles away is reported by neither list", () => {
    const { applied } = patched({ from: 0, to: 2, patch: { bold: null } });
    expect(applied.revisions).toEqual([]);
    expect(applied.folded).toEqual([]);
  });
});
