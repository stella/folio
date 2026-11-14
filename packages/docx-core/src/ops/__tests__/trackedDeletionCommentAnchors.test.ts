/**
 * A tracked deletion of a range takes in a record it covers whole with the
 * comment anchors the record holds, at any depth.
 *
 * Deleting a range directly removes the comment anchors inside it, so
 * accepting the tracked deletion removes every anchor it holds too, those of
 * comment ranges that continue past it included. A planned deletion still
 * leaves a bare anchor standing and deletes around it.
 */
import { describe, expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { planTrackedDeletion } from "../plan";
import {
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const revision = { id: 1, author: "C", date: "2026-03-04T05:00:00Z" };
const newIds = {
  revision: Array.from({ length: 16 }, (_, index) => 303_648 + index),
  control: Array.from({ length: 8 }, (_, index) => 293_648 + index),
};

/**
 * Two content controls: the first holds comment 24's start and "a" with a
 * tab (offsets 0..2), the second holds comment 24's end and "a" (2..3).
 */
const documentOf = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            {
              type: "inlineSdt",
              properties: { sdtType: "richText", id: 1, tag: "clause" },
              content: [
                {
                  type: "hyperlink",
                  rId: "rId9",
                  href: "https://example.org",
                  children: [
                    {
                      type: "inlineWrapper",
                      kind: "bidi",
                      control: "embedding",
                      content: [
                        { type: "commentRangeStart", id: 24 },
                        { type: "commentRangeEnd", id: 1 },
                      ],
                    },
                  ],
                },
                {
                  type: "run",
                  content: [{ type: "text", text: "a" }, { type: "tab" }],
                },
              ],
            },
            {
              type: "inlineSdt",
              properties: { sdtType: "richText", id: 2, tag: "clause" },
              content: [
                {
                  type: "hyperlink",
                  rId: "rId9",
                  href: "https://example.org",
                  children: [
                    {
                      type: "inlineWrapper",
                      kind: "bidi",
                      control: "embedding",
                      content: [{ type: "commentRangeEnd", id: 24 }],
                    },
                  ],
                },
                {
                  type: "hyperlink",
                  rId: "rId9",
                  href: "https://example.org",
                  children: [
                    { type: "run", content: [{ type: "text", text: "a" }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset,
});

const directDeletion = (from: number, to: number): DeleteRangeOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_RANGE,
  from: at(from),
  to: at(to),
});
const trackedDeletion = (from: number, to: number): DeleteRangeOp => ({
  ...directDeletion(from, to),
  revision,
  newIds,
});

const commentEnds = (document: Document): number[] => {
  const found: number[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    if (Reflect.get(value, "type") === "commentRangeEnd") {
      const id: unknown = Reflect.get(value, "id");
      if (typeof id === "number") found.push(id);
    }
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return found;
};

describe("a tracked deletion over a record holding a comment anchor", () => {
  test("a deletion over a content control holding a comment's end accepts to the direct deletion", () => {
    const document = documentOf();
    const tracked = applyDocumentOp(document, trackedDeletion(0, 3));
    if (tracked.isErr()) throw tracked.error;
    const accepted = applyDocumentOp(tracked.value.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.value.revisions,
      decision: REVISION_DECISIONS.ACCEPT,
    });
    if (accepted.isErr()) throw accepted.error;

    // Deleting the range directly removes the control and the end with it.
    const deleted = applyDocumentOp(document, directDeletion(0, 3));
    if (deleted.isErr()) throw deleted.error;
    expect(commentEnds(deleted.value.document)).toEqual([1]);
    expect(accepted.value.document.package.document.content).toEqual(
      deleted.value.document.package.document.content,
    );

    const rejected = applyDocumentOp(tracked.value.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.value.revisions,
      decision: REVISION_DECISIONS.REJECT,
    });
    if (rejected.isErr()) throw rejected.error;
    expect(commentEnds(rejected.value.document)).toEqual([1, 24]);
  });

  test("a deletion beside the record holding the anchor accepts to the direct deletion", () => {
    const document = documentOf();
    const tracked = applyDocumentOp(document, trackedDeletion(0, 2));
    if (tracked.isErr()) throw tracked.error;
    const accepted = applyDocumentOp(tracked.value.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: tracked.value.revisions,
      decision: REVISION_DECISIONS.ACCEPT,
    });
    if (accepted.isErr()) throw accepted.error;
    const deleted = applyDocumentOp(document, directDeletion(0, 2));
    if (deleted.isErr()) throw deleted.error;
    expect(accepted.value.document.package.document.content).toEqual(
      deleted.value.document.package.document.content,
    );
  });

  test("a planned deletion of the same range deletes around the anchor", () => {
    const document = documentOf();
    const planned = planTrackedDeletion(document, {
      from: at(0),
      to: at(3),
      revision,
      newIds,
    });
    if (planned.isErr()) throw planned.error;
    const applied = applyDocumentOps(document, planned.value);
    if (applied.isErr()) throw applied.error;
    expect(commentEnds(applied.value.document)).toEqual([1, 24]);
  });
});
