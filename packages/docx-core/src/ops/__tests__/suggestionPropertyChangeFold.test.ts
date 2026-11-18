/**
 * A paragraph holds one pending property change. A suggestion's stamp folds
 * into it only when the change is that suggestion's own: rejecting the
 * suggestion by its id would otherwise discard another change with it.
 */

import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, type RevisionStamp } from "../types";

const DATE = "2026-01-02T03:04:05Z";

const pending = (suggestionId?: string): Paragraph => ({
  type: "paragraph",
  paraId: "00000002",
  content: [],
  formatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
  propertyChanges: [
    {
      type: "paragraphPropertyChange",
      info: {
        id: 1,
        author: "A",
        ...(suggestionId === undefined ? {} : { suggestionId }),
      },
      previousFormatting: { alignment: "end" },
    },
  ],
});

const documentOf = (second: Paragraph): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
        },
        second,
      ],
    },
  },
});

const stamp = (suggestionId?: string): RevisionStamp => ({
  id: 2,
  author: "S",
  date: DATE,
  ...(suggestionId === undefined ? {} : { suggestionId }),
});

const operations = {
  "a property change": (revision: RevisionStamp) =>
    ({
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "00000002",
      patch: { alignment: "center" },
      revision,
    }) as const,
  "a join into the paragraph": (revision: RevisionStamp) =>
    ({
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      revision,
    }) as const,
};

describe.each(Object.entries(operations))("%s", (_, operation) => {
  test("a suggestion does not take over another author's pending change", () => {
    const tracked = applyDocumentOp(
      documentOf(pending()),
      operation(stamp("s1")),
    );
    expect(tracked.isErr() && tracked.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    );
  });

  test("a suggestion does not take over another suggestion's pending change", () => {
    const tracked = applyDocumentOp(
      documentOf(pending("s0")),
      operation(stamp("s1")),
    );
    expect(tracked.isErr() && tracked.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    );
  });

  test("a suggestion folds into its own pending change", () => {
    const tracked = applyDocumentOp(
      documentOf(pending("s1")),
      operation(stamp("s1")),
    );
    if (tracked.isErr()) throw tracked.error;
    const [, second] = tracked.value.document.package.document.content;
    expect(second).toMatchObject({
      propertyChanges: [
        {
          info: { id: 1, author: "S", suggestionId: "s1" },
          previousFormatting: { alignment: "end" },
        },
      ],
    });
  });

  test("an ordinary tracked change still folds into the pending change", () => {
    const tracked = applyDocumentOp(documentOf(pending()), operation(stamp()));
    if (tracked.isErr()) throw tracked.error;
    const [, second] = tracked.value.document.package.document.content;
    expect(second).toMatchObject({
      propertyChanges: [
        {
          info: { id: 1, author: "S" },
          previousFormatting: { alignment: "end" },
        },
      ],
    });
  });
});
