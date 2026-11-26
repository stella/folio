/**
 * Deleting, tracked, the second half of one's own split of a paragraph one
 * inserted at the story's end takes the split back. That paragraph took over
 * the story's final mark, and with it a pending change from the properties
 * the old last paragraph had; the split moved that mark to the second half.
 * Taking the split back returns the mark, and its pending change, to the
 * paragraph it broke: accepting everything leaves that paragraph's
 * properties, as the direct edits do, and rejecting everything gives the
 * document back. A restyle of the second half after the split is part of
 * what the deletion takes back.
 */
import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

type Mode = "direct" | "tracked-changes";

const LAST = "0000001C";
const TEXT = "Party agreement schedule payment.";

const source = async () => {
  const document = fromMarkdown(
    "First words.\n\nSecond words.\n\nThird words.",
  );
  const last = document.package.document.content.at(-1);
  if (last?.type === "paragraph") last.paraId = LAST;
  const prepared = await prepareCanonicalDocxInput(await createDocx(document));
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Editor",
    session: "canonical",
  });

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operation: FolioDocumentOperation,
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  ).skipped;

const shape = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text, styleId }) => ({ text, styleId }));

const idOf = (reviewer: FolioDocxReviewer, start: string) => {
  const block = reviewer
    .getContent()
    .find(({ text }) => text.startsWith(start));
  if (block === undefined) throw new Error(`No block starts with ${start}`);
  return block.id;
};

/** The edits, and the result once every tracked change is accepted. */
const edit = async (buffer: ArrayBuffer, mode: Mode, restyle: boolean) => {
  const reviewer = await open(buffer);
  expect(
    apply(reviewer, mode, {
      id: "insert",
      type: "insertAfterBlock",
      blockId: LAST,
      text: TEXT,
      styleId: "Heading2",
    }),
  ).toEqual([]);
  expect(
    apply(reviewer, mode, {
      id: "split",
      type: "splitBlock",
      blockId: idOf(reviewer, "Party"),
      offset: 6,
    }),
  ).toEqual([]);
  if (restyle)
    expect(
      apply(reviewer, mode, {
        id: "restyle",
        type: "setBlockParagraphProperties",
        blockId: idOf(reviewer, "agreement"),
        properties: { styleId: "Heading3" },
      }),
    ).toEqual([]);
  expect(
    apply(reviewer, mode, {
      id: "delete",
      type: "deleteBlock",
      blockId: idOf(reviewer, "agreement"),
    }),
  ).toEqual([]);
  return reviewer;
};

describe("a tracked deletion taking back one's own split at the story end", () => {
  for (const restyle of [false, true])
    test(`keeps the broken paragraph's properties (${restyle ? "second half restyled" : "as split"})`, async () => {
      const buffer = await source();
      const direct = shape(await edit(buffer, "direct", restyle));
      expect(direct.at(-1)).toEqual({ text: "Party ", styleId: "Heading2" });

      const tracked = await edit(buffer, "tracked-changes", restyle);
      const rejected = await open(await tracked.toBuffer());
      tracked.acceptAll();
      expect(shape(tracked)).toEqual(direct);

      rejected.rejectAll();
      expect(shape(rejected)).toEqual(shape(await open(buffer)));
    });
});
