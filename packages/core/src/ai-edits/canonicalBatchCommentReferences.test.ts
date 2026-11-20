import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const source = "Alpha beta gamma delta";
const blockId = "10000000";
/** Each word, and the words with the space after them. */
const SPANS = [
  [0, 5],
  [0, 6],
  [6, 10],
  [11, 17],
  [17, 22],
] as const;

const range = ([start, end]: readonly [number, number]) => ({
  type: "textRange" as const,
  story: "main" as const,
  blockId,
  startOffset: start,
  endOffset: end,
  selectedTextHash: hashFolioAIBlockText(source.slice(start, end)),
});

const edits = (span: readonly [number, number]): FolioDocumentOperation[] => [
  { id: "edit", type: "replaceRange", range: range(span), replace: "X" },
  {
    id: "edit",
    type: "formatRange",
    range: range(span),
    formatting: { bold: true },
  },
  { id: "edit", type: "splitBlock", blockId, offset: span[0] },
];

type Mode = "direct" | "tracked-changes";

let buffer: ArrayBuffer | undefined;
const run = async (mode: Mode, batches: FolioDocumentOperation[][]) => {
  buffer ??= await paragraphsDocx([[textRun(source)]]);
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer, {
    session: "canonical",
  });
  const applied: string[] = [];
  for (const operations of batches) {
    const result = reviewer.applyDocumentOperations(
      { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
      { undefinedReferences: "refuse" },
    );
    applied.push(...result.applied.map(({ id }) => id));
  }
  if (mode === "tracked-changes") reviewer.acceptAll();
  return {
    applied: applied.toSorted(),
    texts: reviewer.getContent().map(({ text }) => text),
    comments: reviewer
      .getComments()
      .map(({ text, anchoredText }) => ({ text, anchoredText })),
  };
};

// The batch runs its comments first: an edit beside a comment the batch adds
// lands where it would once that comment stands, one operation at a time.
for (const mode of ["direct", "tracked-changes"] as const)
  test(`${mode}: an edit beside a comment of its batch lands as it would after that comment`, async () => {
    for (const commentSpan of SPANS)
      for (const editSpan of SPANS) {
        if (editSpan[0] < commentSpan[1] && commentSpan[0] < editSpan[1])
          continue;
        for (const edit of edits(editSpan)) {
          if (edit.type === "splitBlock" && edit.offset === 0) continue;
          const comment: FolioDocumentOperation = {
            id: "comment",
            type: "commentOnRange",
            range: range(commentSpan),
            comment: { text: "Check" },
          };
          const context = `${JSON.stringify(commentSpan)} ${edit.type} ${JSON.stringify(editSpan)}`;
          const batched = await run(mode, [[edit, comment]]);
          const sequential = await run(mode, [[comment], [edit]]);
          expect(batched.applied, context).toEqual(["comment", "edit"]);
          expect(batched, context).toEqual(sequential);
        }
      }
  }, 60_000);
