/**
 * Resolving tracked work that builds on pending revisions of its own: the
 * sequences the property in `trackedResolution.property.test.ts` found, each
 * pinned as the smallest one that showed it.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
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

const open = async (bytes?: ArrayBuffer): Promise<FolioDocxReviewer> => {
  if (bytes) {
    return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
  }
  const document = fromMarkdown(["# Service Agreement", ...ORIGINAL.slice(1)].join("\n\n"));
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return FolioDocxReviewer.fromBuffer(
    docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer,
    { author: "Reviewer" },
  );
};

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  open(await reviewer.toBuffer());

const texts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ text }) => text);

const idOf = (reviewer: FolioDocxReviewer, prefix: string): string => {
  const block = reviewer.getContent().find(({ text }) => text.startsWith(prefix));
  if (!block) {
    throw new Error(`no block starts with "${prefix}"`);
  }
  return block.id;
};

const applier =
  (reviewer: FolioDocxReviewer, mode: FolioDocumentOperationMode = "tracked-changes") =>
  (operation: Record<string, unknown>): void => {
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "1", ...operation }],
    } as never);
    expect(result.applied).toHaveLength(1);
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
    apply({ type: "splitBlock", blockId: idOf(reviewer, "Signed"), offset: "Signed in ".length });
    apply({ type: "mergeBlockWithNext", blockId: idOf(reviewer, "two copies"), separator: " " });
    return reviewer;
  };

  test("retracts that insertion: the paragraphs join at once", async () => {
    expect(texts(await build()).slice(-2)).toEqual(["Signed in ", "two copies. Inserted clause."]);
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
    expect(texts(reviewer).slice(-2)).toEqual(["Signed in ", "two copies. Inserted clause."]);
  });

  test("merging the last paragraph into an insertion after it rejects to the original", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({ type: "insertAfterBlock", blockId: idOf(reviewer, "Signed"), text: "Inserted." });
    apply({ type: "mergeBlockWithNext", blockId: idOf(reviewer, "Signed"), separator: " " });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });

  test("merging a split's first half into its deleted second half rejects to the original", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({ type: "splitBlock", blockId: idOf(reviewer, "The Supplier"), offset: 13 });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "delivers") });
    apply({ type: "mergeBlockWithNext", blockId: idOf(reviewer, "The Supplier"), separator: " " });
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(ORIGINAL);
  });
});

describe("a tracked deletion of a paragraph that is a pending insertion", () => {
  for (const mode of ["tracked-changes", "suggested"] as const) {
    for (const [where, anchor] of [
      ["inside the story", "The Buyer"],
      ["at its end", "Signed"],
    ] as const) {
      test(`retracts it ${where} (${mode})`, async () => {
        const reviewer = await open();
        const apply = applier(reviewer, mode);
        apply({ type: "insertAfterBlock", blockId: idOf(reviewer, anchor), text: "Inserted." });
        apply({ type: "deleteBlock", blockId: idOf(reviewer, "Inserted.") });
        expect(texts(reviewer)).toEqual(ORIGINAL);
        expect(reviewer.getChanges()).toEqual([]);
        reviewer.acceptAll();
        expect(texts(await reopen(reviewer))).toEqual(ORIGINAL);
      });
    }
  }

  test("deleting the paragraph a split added a break to keeps its words deleted", async () => {
    const reviewer = await open();
    const apply = applier(reviewer);
    apply({ type: "splitBlock", blockId: idOf(reviewer, "Signed"), offset: "Signed in ".length });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "Signed in") });
    const rejecting = await reopen(reviewer);
    rejecting.rejectAll();
    expect(texts(rejecting)).toEqual(ORIGINAL);
    reviewer.acceptAll();
    expect(texts(reviewer)).toEqual([...ORIGINAL.slice(0, -1), "two copies."]);
  });
});

test("a blank final paragraph deletes like a direct edit and rejects to the blank source", async () => {
  const lastBlockId = (reviewer: FolioDocxReviewer): string => {
    const last = reviewer.getContent().at(-1);
    if (!last) throw new Error("no last block");
    return last.id;
  };
  const source = await open();
  const lastId = idOf(source, "Signed");
  const blanked = source.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "direct",
    operations: [{ id: "blank", type: "replaceBlock", blockId: lastId, text: "" }],
  });
  expect(blanked.applied).toHaveLength(1);
  const blankSource = await source.toBuffer();

  const direct = await open(blankSource);
  const tracked = await open(blankSource);
  const directResult = direct.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "direct",
    operations: [{ id: "delete", type: "deleteBlock", blockId: lastBlockId(direct) }],
  });
  const trackedResult = tracked.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [{ id: "delete", type: "deleteBlock", blockId: lastBlockId(tracked) }],
  });
  expect(directResult.applied).toHaveLength(1);
  expect(trackedResult.applied).toHaveLength(1);
  tracked.acceptAll();
  expect(texts(await reopen(tracked))).toEqual(texts(await reopen(direct)));

  const rejecting = await open(blankSource);
  const rejected = rejecting.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [{ id: "delete", type: "deleteBlock", blockId: lastBlockId(rejecting) }],
  });
  expect(rejected.applied).toHaveLength(1);
  rejecting.rejectAll();
  expect(texts(await reopen(rejecting))).toEqual(texts(await open(blankSource)));
});

describe("inserting after a deleted final paragraph", () => {
  test("rejecting a later insertion after the deleted paragraph restores the source", async () => {
    const reviewer = await open();
    const signedId = idOf(reviewer, "Signed");
    const headingId = idOf(reviewer, "Service Agreement");
    const apply = applier(reviewer);
    apply({ type: "deleteBlock", blockId: signedId });
    apply({ type: "insertAfterBlock", blockId: headingId, text: "Inserted clause." });
    apply({ type: "insertAfterBlock", blockId: signedId, text: "Inserted clause." });
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
    apply({ type: "insertAfterBlock", blockId: headingId, text: "Inserted clause." });
    apply({ type: "insertAfterBlock", blockId: buyerId, text: "Inserted clause." });

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
            const result = reviewer.applyDocumentOperations({
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
            });
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
    apply({ type: "splitBlock", blockId: idOf(reviewer, "The Supplier"), offset: 47 });
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

const anchorsOf = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getComments().map(({ anchoredText }) => anchoredText);

describe("a comment on words a tracked replacement removes", () => {
  const commentThenReplace = async (mode: FolioDocumentOperationMode) => {
    const reviewer = await open();
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

  test("stays one stretch: it reads the same before and after a save", async () => {
    const reviewer = await commentThenReplace("tracked-changes");
    // The removed word, and the new one that takes its place.
    expect(anchorsOf(reviewer)).toEqual(["goodpromptly"]);
    expect(anchorsOf(await reopen(reviewer))).toEqual(anchorsOf(reviewer));
  });

  test("accepted, covers what the direct edit's comment covers", async () => {
    const reviewer = await commentThenReplace("tracked-changes");
    reviewer.acceptAll();
    expect(anchorsOf(reviewer)).toEqual(anchorsOf(await commentThenReplace("direct")));
  });

  test("a replacement right after the comment still continues it", async () => {
    const reviewer = await open();
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
    const anchored = reviewer.getComments().map(({ anchoredText }) => anchoredText);
    expect(anchored).toEqual(["ordershape"]);
    expect((await reopen(reviewer)).getComments().map(({ anchoredText }) => anchoredText)).toEqual(
      anchored,
    );
  });
});

describe("accepting every suggestion", () => {
  test("keeps an insertion after a deleted final paragraph beside a later insertion", async () => {
    const reviewer = await open();
    const signedId = idOf(reviewer, "Signed");
    const headingId = idOf(reviewer, "Service Agreement");
    const apply = applier(reviewer, "suggested");
    apply({ type: "insertAfterBlock", blockId: signedId, text: "Inserted clause." });
    apply({ type: "deleteBlock", blockId: signedId });
    apply({ type: "insertAfterBlock", blockId: headingId, text: "Inserted clause." });

    const direct = await open();
    const directSignedId = idOf(direct, "Signed");
    const directHeadingId = idOf(direct, "Service Agreement");
    const applyDirect = applier(direct, "direct");
    applyDirect({ type: "insertAfterBlock", blockId: directSignedId, text: "Inserted clause." });
    applyDirect({ type: "deleteBlock", blockId: directSignedId });
    applyDirect({ type: "insertAfterBlock", blockId: directHeadingId, text: "Inserted clause." });

    reviewer.acceptAll();
    expect(texts(reviewer)).toEqual(texts(direct));
  });

  test("keeps each suggested paragraph when one of them cannot become a tracked change", async () => {
    const reviewer = await open();
    const apply = applier(reviewer, "suggested");
    apply({ type: "insertAfterBlock", blockId: idOf(reviewer, "This agreement"), text: "Kept." });
    // An insertion after the story's last paragraph, then that paragraph's
    // deletion: the insertion's break can no longer rotate onto it.
    apply({ type: "insertAfterBlock", blockId: idOf(reviewer, "Signed"), text: "Last." });
    apply({ type: "deleteBlock", blockId: idOf(reviewer, "Signed") });
    reviewer.acceptAll();
    const accepted = [...ORIGINAL.slice(0, 2), "Kept.", ...ORIGINAL.slice(2, -1), "Last."];
    expect(texts(reviewer)).toEqual(accepted);
    expect(texts(await reopen(reviewer))).toEqual(accepted);
  });
});
