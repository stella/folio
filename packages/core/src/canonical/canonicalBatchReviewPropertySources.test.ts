/**
 * The tracked batch review reads a paragraph a rejection rebuilds with the
 * properties the original states: list indentation the numbering gives is
 * not restated, so a split of a list item beside a pending merge is no
 * change Reject All would make, and it applies on the live reviewer as on a
 * fork of it.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { FolioDocumentOperation } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const batch = (operations: FolioDocumentOperation[]) => ({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  mode: "tracked-changes" as const,
  operations,
});

test("a split beside a pending merge of list items applies and rejects to the original", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(
      fromMarkdown(
        "Intro words.\n\n- Goods are packed securely\n- Goods are insured in transit\n\nClosing words.",
      ),
    ),
    { author: "Editor", session: "canonical" },
  );
  const item = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === "Goods are packed securely");
  if (item === undefined) throw new Error("the fixture has no list item");
  const original = await reviewer.toBuffer();
  const merged = reviewer.applyDocumentOperations(
    batch([
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: item.id,
        separator: " ",
        mergedParagraphProperties: {
          styleId: null,
          numbering: { kind: "inherit" },
        },
      },
    ]),
    { undefinedReferences: "refuse" },
  );
  expect(merged.skipped).toEqual([]);
  const split = batch([
    { id: "split", type: "splitBlock", blockId: item.id, offset: 10 },
  ]);
  const fork = reviewer.fork();
  const live = reviewer.applyDocumentOperations(split, {
    undefinedReferences: "refuse",
  });
  expect(live.skipped).toEqual([]);
  expect(live.applied.map(({ id }) => id)).toEqual(
    fork
      .applyDocumentOperations(structuredClone(split), {
        undefinedReferences: "refuse",
      })
      .applied.map(({ id }) => id),
  );
  reviewer.rejectAll();
  const texts = (blocks: { text: string }[]) => blocks.map(({ text }) => text);
  expect(texts(reviewer.snapshot().blocks)).toEqual(
    texts(
      (
        await FolioDocxReviewer.fromBuffer(original, { session: "canonical" })
      ).snapshot().blocks,
    ),
  );
});
