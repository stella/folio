import { expect, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const blockId = "10000000";
const operations = [
  {
    id: "replacement",
    type: "replaceInBlock",
    blockId,
    find: "Alpha",
    replace: "Changed",
  },
  {
    id: "insertion",
    type: "insertAfterBlock",
    blockId,
    text: "Added paragraph",
  },
  {
    id: "comment",
    type: "commentOnBlock",
    blockId,
    comment: { text: "Proposed comment" },
  },
] as const satisfies readonly FolioDocumentOperation[];

for (const operation of operations)
  test(`final views refuse a staged ${operation.id} without changing the live reviewer`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await paragraphsDocx([[textRun("Alpha")], [textRun("Beta")]]),
      { session: "canonical" },
    );
    const ordinary = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "ordinary",
            type: "replaceInBlock",
            blockId: "10000001",
            find: "Beta",
            replace: "Updated",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(ordinary.applied).toHaveLength(1);
    expect(ordinary.skipped).toEqual([]);
    const proposal = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    );
    expect(proposal.applied).toHaveLength(1);
    expect(proposal.skipped).toEqual([]);
    const state = () => ({
      document: reviewer.toDocument(),
      blocks: reviewer.getContent(),
      changes: reviewer.getChanges(),
      comments: reviewer.getComments(),
      suggestions: reviewer.exportPendingSuggestions(),
    });
    const before = state();
    for (const read of [
      () => reviewer.readReviewedStory({ view: "final" }),
      () => reviewer.toResolvedDocument("accept"),
    ]) {
      expect(read).toThrow(
        expect.objectContaining({
          reason: "refused",
          ground: { gap: CANONICAL_GAP.publicHeadlessSession },
        }),
      );
      expect(state()).toEqual(before);
    }
    expect(reviewer.acceptSuggestion(operation.id)).toBe(true);
    const final = reviewer.readReviewedStory({ view: "final" });
    expect(final?.changes).toEqual([]);
    const expected = {
      replacement: ["Changed", "Updated"],
      insertion: ["Alpha", "Added paragraph", "Updated"],
      comment: ["Alpha", "Updated"],
    } satisfies Record<(typeof operations)[number]["id"], string[]>;
    expect(final?.snapshot.blocks.map(({ text }) => text)).toEqual(
      expected[operation.id],
    );
  });
