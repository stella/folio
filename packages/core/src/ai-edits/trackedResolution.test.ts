/**
 * Resolving tracked work that builds on pending revisions of its own: the
 * sequences the property in `trackedResolution.property.test.ts` found, each
 * pinned as the smallest one that showed it.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";

const SUPPLIER = "The Supplier delivers the goods on time and in good order.";
const SIGNED = "Signed in two copies.";
const ORIGINAL = [
  "Service Agreement",
  "This agreement is made between the parties named below.",
  SUPPLIER,
  "The Buyer pays each invoice within thirty days.",
  SIGNED,
];

const open = async (
  bytes?: ArrayBuffer,
  session?: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> => {
  if (bytes) {
    return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer", session });
  }
  const document = fromMarkdown(
    ["# Service Agreement", ...ORIGINAL.slice(1)].join("\n\n"),
  );
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  return FolioDocxReviewer.fromBuffer(
    docx.buffer.slice(
      docx.byteOffset,
      docx.byteOffset + docx.byteLength,
    ) as ArrayBuffer,
    { author: "Reviewer", session },
  );
};

const reopen = async (
  reviewer: FolioDocxReviewer,
  session?: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> => open(await reviewer.toBuffer(), session);

const texts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ text }) => text);

/** Each block's kind and text: a restyled paragraph reads differently. */
const kinds = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ kind, text }) => `${kind}: ${text}`);

const ORIGINAL_KINDS = ORIGINAL.map(
  (text, index) => `${index === 0 ? "heading" : "paragraph"}: ${text}`,
);

const idOf = (reviewer: FolioDocxReviewer, prefix: string): string => {
  const block = reviewer
    .getContent()
    .find(({ text }) => text.startsWith(prefix));
  if (!block) {
    throw new Error(`no block starts with "${prefix}"`);
  }
  return block.id;
};

const applier =
  (
    reviewer: FolioDocxReviewer,
    mode: FolioDocumentOperationMode = "tracked-changes",
  ) =>
  (operation: Record<string, unknown>) => {
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [{ id: "1", ...operation }],
      } as never,
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(1);
    return result;
  };

describe("a tracked merge of a break that is itself a pending insertion", () => {
  // Insert after the last paragraph, split that paragraph, and merge the
  // split's second half into the insertion: the merged break is the one the
  // insertion added.
  const build = async (): Promise<FolioDocxReviewer> => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "Signed"),
      text: "Inserted clause.",
    });
    apply({
      type: "splitBlock",
      blockId: idOf(reviewer, "Signed"),
      offset: "Signed in ".length,
    });
    apply({
      type: "mergeBlockWithNext",
      blockId: idOf(reviewer, "two copies"),
      separator: " ",
    });
    return reviewer;
  };

  test("retracts that insertion: the paragraphs join at once", async () => {
    expect(texts(await build()).slice(-2)).toEqual([
      "Signed in ",
      "two copies. Inserted clause.",
    ]);
  });

  test("rejecting everything gives the original back", async () => {
    const reviewer = await build();
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
    expect(texts(await reopen(reviewer))).toEqual(ORIGINAL);
  });

  test("accepting everything keeps the split and the merge", async () => {
    const reviewer = await build();
    reviewer.acceptAll();
    expect(texts(reviewer).slice(-2)).toEqual([
      "Signed in ",
      "two copies. Inserted clause.",
    ]);
  });

  test("merging the last paragraph into an insertion after it rejects to the original", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "Signed"),
      text: "Inserted.",
    });
    apply({
      type: "mergeBlockWithNext",
      blockId: idOf(reviewer, "Signed"),
      separator: " ",
    });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });

  test("merging a split's first half into its deleted second half rejects to the original", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({
      type: "splitBlock",
      blockId: idOf(reviewer, "The Supplier"),
      offset: 13,
    });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "delivers") });
    apply({
      type: "mergeBlockWithNext",
      blockId: idOf(reviewer, "The Supplier"),
      separator: " ",
    });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });

  // Seed 386876854: the split's first half reads as the heading merged into
  // it; merging that half on hands the heading to the paragraph left, which
  // must go back to a body paragraph when rejected.
  describe("merging a split half that reads as the heading merged into it", () => {
    const restyledHalfMerged = async (): Promise<FolioDocxReviewer> => {
      const reviewer = await open();
      const apply = applier(reviewer);
      apply({
        type: "splitBlock",
        blockId: idOf(reviewer, "This agreement"),
        offset: 5,
      });
      apply({
        type: "mergeBlockWithNext",
        blockId: idOf(reviewer, "Service"),
        separator: " ",
      });
      apply({
        type: "mergeBlockWithNext",
        blockId: idOf(reviewer, "This "),
        separator: " ",
      });
      return reviewer;
    };

    test("rejects to the original body paragraph, live and saved", async () => {
      const reviewer = await restyledHalfMerged();
      const saved = await reopen(reviewer);
      reviewer.rejectAll();
      expect(kinds(reviewer)).toEqual(ORIGINAL_KINDS);
      expect(kinds(await reopen(reviewer))).toEqual(ORIGINAL_KINDS);
      saved.rejectAll();
      expect(kinds(saved)).toEqual(ORIGINAL_KINDS);
    });

    test("accepts to the heading, live and saved", async () => {
      const reviewer = await restyledHalfMerged();
      const saved = await reopen(reviewer);
      const expected = [
        "heading: Service Agreement This  agreement is made between the parties named below.",
        ...ORIGINAL_KINDS.slice(2),
      ];
      reviewer.acceptAll();
      expect(kinds(reviewer)).toEqual(expected);
      expect(kinds(await reopen(reviewer))).toEqual(expected);
      saved.acceptAll();
      expect(kinds(saved)).toEqual(expected);
    });
  });
});

describe("a tracked deletion of a paragraph that is a pending insertion", () => {
  // Deleting one's own inserted paragraph takes the insertion back: the
  // document is the one before it, ids and paragraph properties included,
  // live, saved, resolved and undone.
  for (const session of ["canonical", "legacy"] as const) {
    for (const mode of ["tracked-changes", "suggested"] as const) {
      for (const [where, anchor] of [
        ["inside the story", "The Buyer"],
        ["at its end", "Signed"],
      ] as const) {
        for (const styleId of [undefined, "Heading2"] as const) {
          test(`retracts it ${where}${styleId === undefined ? "" : " restyled"} (${session}, ${mode})`, async () => {
            const reviewer = await open(undefined, session);
            const original = (
              await open(await reviewer.toBuffer(), session)
            ).getContent();
            const apply = applier(reviewer, mode);
            const inserted = apply({
              type: "insertAfterBlock",
              blockId: idOf(reviewer, anchor),
              text: "Inserted.",
              ...(styleId === undefined ? {} : { styleId }),
            });
            // A suggestion takes back only its own records: the deletion
            // joins the insertion's suggestion.
            const suggestionId = inserted.applied.at(0)?.suggestionId;
            const deleted = apply({
              id: "2",
              type: "deleteBlock",
              blockId: idOf(reviewer, "Inserted."),
              ...(suggestionId === undefined ? {} : { suggestionId }),
            });
            expect(reviewer.getContent()).toEqual(original);
            expect(reviewer.getChanges()).toEqual([]);
            const saved = await reopen(reviewer, session);
            expect(saved.getContent()).toEqual(original);
            expect(saved.getChanges()).toEqual([]);
            saved.rejectAll();
            expect(saved.getContent()).toEqual(original);
            const accepted = await reopen(reviewer, session);
            accepted.acceptAll();
            expect(accepted.getContent()).toEqual(original);
            const undoHandle = deleted.undoHandle;
            if (undoHandle === null) throw new Error("no undo handle");
            expect(reviewer.undoDocumentOperations(undoHandle).status).toBe(
              "undone",
            );
            expect(texts(reviewer)).toContain("Inserted.");
          });
        }
      }
    }
  }

  // A later suggestion is not the one that inserted the paragraph: on the
  // canonical session its deletion stays beside the insertion and the two
  // suggestions stay independent. Rejecting the deletion alone gives the
  // paragraph back, rejecting the insertion alone leaves the document before
  // both, and accepting both, in either order, removes the paragraph. The
  // legacy session takes the insertion back outright.
  const suggestedThenDeleted = async (session: HeadlessReviewerSession) => {
    const reviewer = await open(undefined, session);
    const original = (
      await open(await reviewer.toBuffer(), session)
    ).getContent();
    const apply = applier(reviewer, "suggested");
    const inserting = apply({
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "The Buyer"),
      text: "Inserted.",
    }).applied.at(0)?.suggestionId;
    const deleting = apply({
      id: "2",
      type: "deleteBlock",
      blockId: idOf(reviewer, "Inserted."),
    }).applied.at(0)?.suggestionId;
    if (inserting === undefined || deleting === undefined)
      throw new Error("no suggestion");
    return { reviewer, original, inserting, deleting };
  };

  test("a later suggestion deleting it is recorded beside it (canonical)", async () => {
    const { reviewer, original } = await suggestedThenDeleted("canonical");
    expect(reviewer.getChanges().map(({ type }) => type)).toEqual([
      "paragraphMarkInserted",
      "paragraphMarkDeleted",
      "insertion",
      "deletion",
    ]);
    const accepted = await reopen(reviewer, "canonical");
    accepted.acceptAll();
    expect(accepted.getContent()).toEqual(original);
    const rejected = await reopen(reviewer, "canonical");
    rejected.rejectAll();
    expect(rejected.getContent()).toEqual(original);
  });

  test("rejecting the later suggestion alone gives the paragraph back (canonical)", async () => {
    const { reviewer, deleting } = await suggestedThenDeleted("canonical");
    expect(reviewer.rejectSuggestion(deleting)).toBe(true);
    expect(texts(reviewer)).toContain("Inserted.");
    expect(reviewer.getChanges().map(({ type }) => type)).toEqual([
      "paragraphMarkInserted",
      "insertion",
    ]);
  });

  test("rejecting the inserting suggestion alone leaves the document before both (canonical)", async () => {
    const { reviewer, original, inserting, deleting } =
      await suggestedThenDeleted("canonical");
    expect(reviewer.rejectSuggestion(inserting)).toBe(true);
    expect(reviewer.getContent()).toEqual(original);
    expect(reviewer.getChanges()).toEqual([]);
    expect(reviewer.acceptSuggestion(deleting)).toBe(false);
  });

  for (const order of ["inserting first", "deleting first"] as const) {
    test(`accepting both suggestions, ${order}, removes the paragraph (canonical)`, async () => {
      const { reviewer, original, inserting, deleting } =
        await suggestedThenDeleted("canonical");
      const sequence =
        order === "inserting first"
          ? [inserting, deleting]
          : [deleting, inserting];
      for (const suggestionId of sequence)
        expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
      reviewer.acceptAll();
      expect(reviewer.getContent()).toEqual(original);
      expect(reviewer.getChanges()).toEqual([]);
    });
  }

  test("a later suggestion deleting it takes it back (legacy)", async () => {
    const { reviewer, original, deleting } =
      await suggestedThenDeleted("legacy");
    expect(reviewer.getContent()).toEqual(original);
    expect(reviewer.getChanges()).toEqual([]);
    expect(reviewer.rejectSuggestion(deleting)).toBe(false);
  });

  // A comment the inserted paragraph alone carries retires with it, as a
  // direct deletion of a commented paragraph retires its comment; the result
  // reports it. The legacy session refuses the deletion.
  for (const [where, anchor] of [
    ["inside the story", "The Buyer"],
    ["at its end", "Signed"],
  ] as const) {
    for (const session of ["canonical", "legacy"] as const) {
      test(`retracting it ${where} retires its comment (${session})`, async () => {
        const reviewer = await open(undefined, session);
        const original = (
          await open(await reviewer.toBuffer(), session)
        ).getContent();
        const commentId = applier(reviewer)({
          type: "insertAfterBlock",
          blockId: idOf(reviewer, anchor),
          text: "Inserted.",
          comment: { text: "Why?" },
        }).applied.at(0)?.commentId;
        if (commentId === undefined) throw new Error("no comment");
        const commented = reviewer.getContent();
        const deleted = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [
              {
                id: "2",
                type: "deleteBlock",
                blockId: idOf(reviewer, "Inserted."),
              },
            ],
          },
          { undefinedReferences: "refuse" },
        );
        if (session === "legacy") {
          expect(deleted.applied).toEqual([]);
          expect(deleted.skipped.map(({ message }) => message)).toEqual([
            "Accept All would lose a comment.",
          ]);
          expect(reviewer.getContent()).toEqual(commented);
          return;
        }
        expect(deleted.applied.at(0)?.retiredCommentIds).toEqual([commentId]);
        expect(reviewer.getContent()).toEqual(original);
        expect(reviewer.getChanges()).toEqual([]);
        expect(reviewer.getComments()).toEqual([]);
        const saved = await reopen(reviewer, session);
        expect(saved.getContent()).toEqual(original);
        expect(saved.getComments()).toEqual([]);
        const undoHandle = deleted.undoHandle;
        if (undoHandle === null) throw new Error("no undo handle");
        expect(reviewer.undoDocumentOperations(undoHandle).status).toBe(
          "undone",
        );
        expect(reviewer.getComments().map(({ id }) => id)).toEqual([commentId]);
      });
    }
  }

  test("deleting the paragraph a split added a break to keeps its words deleted", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({
      type: "splitBlock",
      blockId: idOf(reviewer, "Signed"),
      offset: "Signed in ".length,
    });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "Signed in") });
    const rejecting = await reopen(reviewer);
    rejecting.rejectAll();
    expect(texts(rejecting)).toEqual(ORIGINAL);
    reviewer.acceptAll();
    expect(texts(reviewer)).toEqual([...ORIGINAL.slice(0, -1), "two copies."]);
  });

  for (const mode of ["tracked-changes", "suggested"] as const) {
    test(`its words run into a restyled insertion and reject with their own style (${mode})`, async () => {
      // Inserting after the last paragraph gives that paragraph the added
      // break; deleting it runs its words into a paragraph inserted after it
      // in the same batch, whose break is added too.
      const reviewer = await open();
      const signed = idOf(reviewer, "Signed");
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "1",
              type: "insertAfterBlock",
              blockId: signed,
              text: "Inserted.",
              styleId: "Heading2",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            { id: "2", type: "deleteBlock", blockId: signed },
            {
              id: "3",
              type: "insertAfterBlock",
              blockId: signed,
              text: "Added.",
              styleId: "Heading2",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(2);
      const styles = (current: FolioDocxReviewer) =>
        current
          .getContent()
          .map(({ text, styleId }) => [text, styleId ?? null]);
      const original = styles(await open());
      const saved = await reopen(reviewer);
      reviewer.rejectAll();
      expect(styles(reviewer)).toEqual(original);
      saved.rejectAll();
      expect(styles(saved)).toEqual(original);
    });
  }
});

test("a blank final paragraph deletes like a direct edit and rejects to the blank source", async () => {
  const lastBlockId = (reviewer: FolioDocxReviewer): string => {
    const last = reviewer.getContent().at(-1);
    if (!last) throw new Error("no last block");
    return last.id;
  };
  const source = await open();
  const lastId = idOf(source, "Signed");
  const blanked = source.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        { id: "blank", type: "replaceBlock", blockId: lastId, text: "" },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(blanked.applied).toHaveLength(1);
  const blankSource = await source.toBuffer();

  const direct = await open(blankSource);
  const tracked = await open(blankSource);
  const directResult = direct.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        { id: "delete", type: "deleteBlock", blockId: lastBlockId(direct) },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const trackedResult = tracked.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "delete", type: "deleteBlock", blockId: lastBlockId(tracked) },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(directResult.applied).toHaveLength(1);
  expect(trackedResult.applied).toHaveLength(1);
  tracked.acceptAll();
  expect(texts(await reopen(tracked))).toEqual(texts(await reopen(direct)));

  const rejecting = await open(blankSource);
  const rejected = rejecting.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "delete", type: "deleteBlock", blockId: lastBlockId(rejecting) },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(rejected.applied).toHaveLength(1);
  rejecting.rejectAll();
  expect(texts(await reopen(rejecting))).toEqual(
    texts(await open(blankSource)),
  );
});

describe("inserting after a deleted final paragraph", () => {
  // The canonical session refuses a tracked table right after a paragraph
  // whose break the final deletion retired, where accepting would run that
  // paragraph's words into the table's first cell; the legacy session places
  // the table after the deleted paragraph. Everything else resolves like
  // direct editing on both.
  const TABLE_AFTER_RETIRED_BREAK = {
    reason: "pendingParagraphMarkDeletion",
    message:
      "Accepting the pending deletion of this paragraph's break would run its words into the inserted table. Resolve that change first, or insert the table elsewhere.",
  };

  for (const session of ["canonical", "legacy"] as const)
    for (const { deleted, anchor } of [
      { deleted: ["Signed"], anchor: "The Buyer" },
      { deleted: ["Signed", "The Buyer"], anchor: "The Supplier" },
    ] as const) {
      test(`${deleted.length} deleted final paragraph(s) before a table (${session})`, async () => {
        const results: string[][] = [];
        for (const mode of ["direct", "tracked-changes"] as const) {
          const reviewer = await open(undefined, session);
          const apply = applier(reviewer, mode);
          for (const target of deleted) {
            apply({ type: "deleteBlock", blockId: idOf(reviewer, target) });
          }
          const table = {
            type: "insertTable",
            blockId: idOf(reviewer, anchor),
            rows: [["Term", "Value"]],
          };
          if (mode === "tracked-changes" && session === "canonical") {
            const before = reviewer.getContent();
            const result = reviewer.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                operations: [{ id: "table", ...table }],
              },
              { undefinedReferences: "refuse" },
            );
            expect(result.applied).toEqual([]);
            expect(result.skipped).toEqual([
              { id: "table", ...TABLE_AFTER_RETIRED_BREAK },
            ]);
            expect(reviewer.getContent()).toEqual(before);
            return;
          }
          apply(table);
          if (mode === "tracked-changes") {
            const rejecting = await reopen(reviewer, session);
            rejecting.rejectAll();
            expect(texts(await reopen(rejecting, session))).toEqual(ORIGINAL);
            reviewer.acceptAll();
          }
          results.push(texts(await reopen(reviewer, session)));
        }
        expect(results[1]).toEqual(results[0]);
      });
    }

  // In one batch, a direct deletion that joins the anchor into the closing
  // paragraph leaves the table after the paragraph that now holds it.
  for (const session of ["canonical", "legacy"] as const)
    for (const { deleted, anchor } of [
      { deleted: ["Signed"], anchor: "The Buyer" },
      { deleted: ["Signed", "The Buyer"], anchor: "The Supplier" },
    ] as const) {
      test(`${deleted.length} final deletion(s) and a table in one batch (${session})`, async () => {
        const results: string[][] = [];
        for (const mode of ["direct", "tracked-changes"] as const) {
          const reviewer = await open(undefined, session);
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [
                ...deleted.map((target, index) => ({
                  id: `delete-${index}`,
                  type: "deleteBlock" as const,
                  blockId: idOf(reviewer, target),
                })),
                {
                  id: "table",
                  type: "insertTable",
                  blockId: idOf(reviewer, anchor),
                  rows: [["Term", "Value"]],
                },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied).toHaveLength(deleted.length + 1);
          if (mode === "tracked-changes") {
            const rejecting = await reopen(reviewer, session);
            rejecting.rejectAll();
            expect(texts(await reopen(rejecting, session))).toEqual(ORIGINAL);
            reviewer.acceptAll();
          }
          results.push(texts(await reopen(reviewer, session)));
        }
        expect(results[1]).toEqual(results[0]);
        // The table lands before the deleted closing paragraph, which then
        // stays emptied: the body still ends with a paragraph.
        expect(results[0]?.slice(-3)).toEqual(["Term", "Value", ""]);
      });
    }

  test("rejecting a later insertion after the deleted paragraph restores the source", async () => {
    const reviewer = await open();
    const signedId = idOf(reviewer, "Signed");
    const headingId = idOf(reviewer, "Service Agreement");
    const apply = applier(reviewer);
    apply({ type: "deleteBlock", blockId: signedId });
    apply({
      type: "insertAfterBlock",
      blockId: headingId,
      text: "Inserted clause.",
    });
    apply({
      type: "insertAfterBlock",
      blockId: signedId,
      text: "Inserted clause.",
    });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });

  test("accepting keeps the insertion on its own paragraph", async () => {
    const reviewer = await open();
    const signedId = idOf(reviewer, "Signed");
    const headingId = idOf(reviewer, "Service Agreement");
    const buyerId = idOf(reviewer, "The Buyer");
    const apply = applier(reviewer);
    apply({ type: "deleteBlock", blockId: signedId });
    apply({
      type: "insertAfterBlock",
      blockId: headingId,
      text: "Inserted clause.",
    });
    apply({
      type: "insertAfterBlock",
      blockId: buyerId,
      text: "Inserted clause.",
    });

    reviewer.acceptAll();
    expect(texts(reviewer).slice(-2)).toEqual([
      "The Buyer pays each invoice within thirty days.",
      "Inserted clause.",
    ]);
    expect(texts(await reopen(reviewer))).toEqual(texts(reviewer));
  });
});

describe("successive trailing paragraph deletions", () => {
  test("an earlier insertion keeps the accepted trailing deletion chain compact", async () => {
    const runSequence = async (mode: FolioDocumentOperationMode) => {
      const reviewer = await open();
      const apply = applier(reviewer, mode);
      apply({ type: "deleteBlock", blockId: idOf(reviewer, "Signed") });
      apply({ type: "deleteBlock", blockId: idOf(reviewer, "The Buyer") });
      apply({
        type: "insertAfterBlock",
        blockId: idOf(reviewer, "Service Agreement"),
        text: "Inserted clause.",
      });
      return reviewer;
    };

    const direct = await runSequence("direct");
    const tracked = await runSequence("tracked-changes");
    tracked.acceptAll();
    expect(texts(tracked)).toEqual(texts(direct));
  });

  for (const count of [2, 3, 4]) {
    for (const batching of ["separate", "together"] as const) {
      test(`${count} adjacent deletions ${batching} resolve like direct edits`, async () => {
        const targets = ORIGINAL.slice(-count);
        const runChain = async (mode: FolioDocumentOperationMode) => {
          const reviewer = await open();
          if (batching === "separate") {
            const apply = applier(reviewer, mode);
            for (const text of targets) {
              apply({ type: "deleteBlock", blockId: idOf(reviewer, text) });
            }
          } else {
            const result = reviewer.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                operations: targets.map(
                  (text, index) =>
                    ({
                      id: String(index),
                      type: "deleteBlock",
                      blockId: idOf(reviewer, text),
                    }) as const,
                ),
              },
              { undefinedReferences: "refuse" },
            );
            expect(result.applied).toHaveLength(count);
          }
          return reviewer;
        };

        const direct = await runChain("direct");
        const tracked = await runChain("tracked-changes");
        const rejecting = await reopen(tracked);
        rejecting.rejectAll();
        expect(texts(rejecting)).toEqual(ORIGINAL);

        tracked.acceptAll();
        expect(texts(tracked)).toEqual(texts(direct));
        expect(texts(await reopen(tracked))).toEqual(texts(direct));
      });
    }
  }
});

describe("rejecting a split with an inserted table after its first half", () => {
  test("joins the halves again once the table is gone", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({
      type: "splitBlock",
      blockId: idOf(reviewer, "The Supplier"),
      offset: 47,
    });
    apply({
      type: "insertTable",
      blockId: idOf(reviewer, "The Supplier"),
      rows: [["Term", "Value"]],
    });
    expect(texts(reviewer)).toContain("Term");
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });
});

test("accepting a split paragraph deletion before a table matches direct editing", async () => {
  const results: string[][] = [];
  for (const mode of ["direct", "tracked-changes"] as const) {
    const reviewer = await open();
    applier(
      reviewer,
      mode,
    )({
      type: "splitBlock",
      blockId: idOf(reviewer, "The Supplier"),
      offset: 20,
    });
    const blockId = idOf(reviewer, "The Supplier");
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [
          { id: "delete", type: "deleteBlock", blockId },
          {
            id: "table",
            type: "insertTable",
            blockId,
            rows: [["Term", "Value"]],
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(2);
    expect(result.skipped).toEqual([]);
    const saved = await reopen(reviewer);
    if (mode === "tracked-changes") {
      const rejecting = await reopen(saved);
      rejecting.rejectAll();
      expect(texts(await reopen(rejecting))).toEqual(ORIGINAL);
      saved.acceptAll();
    }
    results.push(texts(await reopen(saved)));
  }
  expect(results[1]).toEqual(results[0]);
});

test("clearing a split paragraph's text before a table keeps its break", async () => {
  const reviewer = await open();
  applier(reviewer)({
    type: "splitBlock",
    blockId: idOf(reviewer, "The Supplier"),
    offset: 20,
  });
  const blockId = idOf(reviewer, "The Supplier");
  applier(reviewer)({ type: "replaceBlock", blockId, text: "" });
  applier(reviewer)({
    type: "insertTable",
    blockId,
    rows: [["Term", "Value"]],
  });
  reviewer.acceptAll();
  expect(texts(await reopen(reviewer))).toContain("");
});

const anchorsOf = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getComments().map(({ anchoredText }) => anchoredText);

describe("a comment on words a tracked replacement removes", () => {
  // The canonical session keeps the comment on the deleted words only, so
  // accepting the replacement retires it, as a direct replacement covering
  // the whole comment deletes it. The legacy session extends it over the
  // inserted replacement.
  const commentThenReplace = async (
    mode: FolioDocumentOperationMode,
    session: HeadlessReviewerSession,
  ) => {
    const reviewer = await open(undefined, session);
    const start = SUPPLIER.indexOf("good order");
    applier(reviewer)({
      type: "commentOnRange",
      range: createFolioAITextRangeHandle({
        blockId: idOf(reviewer, "The Supplier"),
        text: SUPPLIER,
        startOffset: start,
        endOffset: start + "good".length,
      }),
      comment: { text: "Which standard?" },
    });
    applier(
      reviewer,
      mode,
    )({
      type: "replaceBlock",
      blockId: idOf(reviewer, "The Supplier"),
      text: "The Supplier delivers promptly.",
    });
    return reviewer;
  };

  for (const [session, replaced, continued] of [
    ["canonical", "good", "order"],
    ["legacy", "good promptly.", "ordershape"],
  ] as const) {
    test(`stays one stretch: it reads the same before and after a save (${session})`, async () => {
      const reviewer = await commentThenReplace("tracked-changes", session);
      expect(anchorsOf(reviewer)).toEqual([replaced]);
      expect(anchorsOf(await reopen(reviewer, session))).toEqual(
        anchorsOf(reviewer),
      );
    });

    test(`accepted, covers what the direct edit's comment covers (${session})`, async () => {
      const reviewer = await commentThenReplace("tracked-changes", session);
      reviewer.acceptAll();
      expect(anchorsOf(reviewer)).toEqual(
        anchorsOf(await commentThenReplace("direct", session)),
      );
    });

    test(`a replacement right after the comment (${session})`, async () => {
      const reviewer = await open(undefined, session);
      const apply = applier(reviewer);
      const start = SUPPLIER.indexOf("order");
      apply({
        type: "commentOnRange",
        range: createFolioAITextRangeHandle({
          blockId: idOf(reviewer, "The Supplier"),
          text: SUPPLIER,
          startOffset: start,
          endOffset: start + "order".length,
        }),
        comment: { text: "Which standard?" },
      });
      apply({
        type: "replaceInBlock",
        blockId: idOf(reviewer, "The Supplier"),
        find: "order",
        replace: "shape",
      });
      const anchored = reviewer
        .getComments()
        .map(({ anchoredText }) => anchoredText);
      expect(anchored).toEqual([continued]);
      expect(
        (await reopen(reviewer, session))
          .getComments()
          .map(({ anchoredText }) => anchoredText),
      ).toEqual(anchored);
    });
  }
});

describe("resolving suggested paragraphs", () => {
  test("keeps two insertions staged around a deleted final paragraph", async () => {
    const reviewer = await open();
    const signedId = idOf(reviewer, "Signed");
    const headingId = idOf(reviewer, "Service Agreement");
    const apply = applier(reviewer, "suggested");
    apply({
      type: "insertAfterBlock",
      blockId: signedId,
      text: "Inserted clause.",
      suggestionId: "last-insertion",
    });
    apply({
      type: "deleteBlock",
      blockId: signedId,
      suggestionId: "signed-deletion",
    });
    apply({
      type: "insertAfterBlock",
      blockId: headingId,
      text: "Inserted clause.",
      suggestionId: "heading-insertion",
    });

    const direct = await open();
    const directSignedId = idOf(direct, "Signed");
    const directHeadingId = idOf(direct, "Service Agreement");
    const applyDirect = applier(direct, "direct");
    applyDirect({
      type: "insertAfterBlock",
      blockId: directSignedId,
      text: "Inserted clause.",
    });
    applyDirect({ type: "deleteBlock", blockId: directSignedId });
    applyDirect({
      type: "insertAfterBlock",
      blockId: directHeadingId,
      text: "Inserted clause.",
    });

    const staged = texts(reviewer);
    expect(staged.filter((text) => text !== "")).toEqual(texts(direct));
    expect(reviewer.acceptAll().resolved).toBe(0);
    expect(texts(reviewer)).toEqual(staged);
    expect(texts(await reopen(reviewer))).toEqual(ORIGINAL);
  });

  test("keeps each suggested paragraph when one of them cannot become a tracked change", async () => {
    const reviewer = await open();
    const apply = applier(reviewer, "suggested");
    apply({
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "This agreement"),
      text: "Kept.",
    });
    // An insertion after the story's last paragraph, then that paragraph's
    // deletion: the insertion's break can no longer rotate onto it.
    apply({
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "Signed"),
      text: "Last.",
    });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "Signed") });
    const staged = texts(reviewer);
    expect(staged).toEqual([
      ...ORIGINAL.slice(0, 2),
      "Kept.",
      ...ORIGINAL.slice(2, -1),
      "",
      "Last.",
    ]);
    expect(reviewer.acceptAll().resolved).toBe(0);
    expect(texts(reviewer)).toEqual(staged);
    expect(texts(await reopen(reviewer))).toEqual(ORIGINAL);
  });
});
