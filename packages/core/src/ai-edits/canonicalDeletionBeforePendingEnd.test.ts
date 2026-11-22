import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const MODES = ["suggested", "tracked-changes"] as const;
type Mode = (typeof MODES)[number];

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: readonly FolioDocumentOperation[],
) => {
  const { skipped } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [...operations],
    },
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
};

/** Four paragraphs, the last one's deletion pending, and their ids. */
const openWithPendingEnd = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await paragraphsDocx(
      ["Title", "Alpha", "Bravo", "Charlie"].map((text) => [textRun(text)]),
    ),
    { session: "canonical" },
  );
  const ids = reviewer.getContent().map(({ id }) => id);
  apply(reviewer, "tracked-changes", [
    { id: "pending", type: "deleteBlock", blockId: ids[3] ?? "" },
  ]);
  return { reviewer, ids };
};

/** What a review shows while the changes pend, and what accepting them leaves. */
const outcome = (reviewer: FolioDocxReviewer) => {
  const shown = () =>
    reviewer
      .getContent()
      .map(({ kind, displayLabel, text }) => ({ kind, displayLabel, text }));
  const pending = shown();
  for (const id of new Set(
    reviewer.exportPendingSuggestions().map(({ suggestionId }) => suggestionId),
  ))
    expect(reviewer.acceptSuggestion(id)).toBe(true);
  reviewer.acceptAll();
  return { pending, accepted: shown() };
};

// The deletions leave only a paragraph a pending deletion takes after them:
// they retire the inserted paragraph's break, not the title's.
for (const mode of MODES)
  for (const numbered of [false, true])
    for (const order of [
      [0, 1, 2],
      [0, 2, 1],
      [1, 0, 2],
      [1, 2, 0],
      [2, 0, 1],
      [2, 1, 0],
    ])
      test(`${mode}: an insertion before deletions that reach a pending end, ${numbered ? "numbered" : "plain"}, in order ${order.join("")}`, async () => {
        const batch = await openWithPendingEnd();
        const alone = await openWithPendingEnd();
        const [title = "", first = "", second = ""] = batch.ids;
        const operations: readonly FolioDocumentOperation[] = [
          { id: "first", type: "deleteBlock", blockId: first },
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: title,
            text: "Payment.",
            ...(numbered && {
              numbering: { kind: "newList", format: "numbered" },
            }),
          },
          { id: "second", type: "deleteBlock", blockId: second },
        ];
        const ordered = order.flatMap((index) => operations[index] ?? []);
        apply(batch.reviewer, mode, ordered);
        for (const operation of ordered)
          apply(alone.reviewer, mode, [operation]);

        const result = outcome(batch.reviewer);
        expect(result.accepted.map(({ text }) => text)).toEqual([
          "Title",
          "Payment.",
        ]);
        // Once a deletion sent before it has retired the title's break, an
        // insertion sent last goes after that join: the batch reads the
        // title as it stood before.
        if (order[2] !== 1) expect(result).toEqual(outcome(alone.reviewer));
      });
