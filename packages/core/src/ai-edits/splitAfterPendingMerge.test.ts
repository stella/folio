/**
 * A tracked split of the paragraph a pending merge runs into, where the first
 * half holds only words an earlier change inserted. The batch-end pass that
 * moves an old deleted break past a paragraph this batch inserted took that
 * half for one: it moved the deleted break onto the half and the split's added
 * break onto the merging paragraph, so accepting kept the merging paragraph
 * apart and ran the half back into the rest, losing the split. One operation
 * at a time after deleting the merging paragraph, the split stood.
 *
 * The split's first half, which ends at the mark the split adds, starts a
 * paragraph of its own once the merge is accepted: it reads with the merged
 * paragraph's properties, as the merged paragraph split directly does.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

type Operation = Parameters<
  FolioDocxReviewer["applyDocumentOperations"]
>[0]["operations"][number];

const buildDocument = (): Document => ({
  package: {
    document: {
      content: [paragraph("Alpha words."), paragraph("Beta words here.")],
    },
  },
});

const tracked = (reviewer: FolioDocxReviewer, operation: Operation) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
};

const idOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const id = reviewer
    .getContent()
    .find((block) => block.text.startsWith(text))?.id;
  if (id === undefined) throw new Error(`no block starts with "${text}"`);
  return id;
};

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map((block) => block.text.trimEnd());

test.each(SESSIONS)(
  "a tracked split of inserted words a pending merge runs into survives accepting (%s)",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(buildDocument()),
      { author: "Editor", session },
    );
    // Inserted words opening "Beta words here.", under that paragraph's own break.
    tracked(reviewer, {
      id: "insert",
      type: "insertBeforeBlock",
      blockId: idOf(reviewer, "Beta"),
      text: "Inserted clause.",
    });
    tracked(reviewer, {
      id: "join-inserted",
      type: "mergeBlockWithNext",
      blockId: idOf(reviewer, "Inserted clause."),
      separator: " ",
    });
    expect(texts(reviewer)).toEqual([
      "Alpha words.",
      "Inserted clause. Beta words here.",
    ]);
    // A pending merge running "Alpha words." into it.
    tracked(reviewer, {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: idOf(reviewer, "Alpha"),
      separator: " ",
    });
    tracked(reviewer, {
      id: "split",
      type: "splitBlock",
      blockId: idOf(reviewer, "Inserted clause."),
      offset: "Inserted clause. ".length,
    });

    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    for (const document of [reviewer, reopened]) {
      document.acceptAll();
      expect(texts(document)).toEqual([
        "Alpha words. Inserted clause.",
        "Beta words here.",
      ]);
    }
  },
);

describe.each(SESSIONS)(
  "a tracked split of a paragraph a pending merge joins on (%s)",
  (session) => {
    const FIXTURE = "1. Numbered clause one\n\nUnnumbered body text.";
    const open = async () =>
      FolioDocxReviewer.fromBuffer(
        (await ensureParaIds(await createDocx(fromMarkdown(FIXTURE)))).docx,
        { author: "Test", session },
      );
    const edit = async (
      mode: "direct" | "tracked-changes",
      firstParagraphProperties?: { alignment: "center" },
    ) => {
      const reviewer = await open();
      const item = reviewer.snapshot().blocks.at(0);
      if (!item) throw new Error("Missing fixture");
      let currentBlockId = item.id;
      for (const operation of [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: item.id,
          separator: " ",
          mergedParagraphProperties: { numbering: { kind: "none" } },
        },
        {
          id: "split",
          type: "splitBlock",
          blockId: item.id,
          offset: 9,
          ...(firstParagraphProperties && { firstParagraphProperties }),
        },
      ] satisfies FolioDocumentOperation[]) {
        const result = reviewer.applyOperations(
          [{ ...operation, blockId: currentBlockId }],
          { mode },
        );
        expect(result.skipped).toEqual([]);
        for (const { paragraphIdentity } of result.applied) {
          if (
            paragraphIdentity?.type === "joined" &&
            paragraphIdentity.retiredBlockId === currentBlockId
          )
            currentBlockId = paragraphIdentity.blockId;
        }
      }
      return reviewer;
    };
    const read = (reviewer: FolioDocxReviewer) =>
      reviewer
        .snapshot()
        .blocks.map(({ text, kind, listReference, alignment }) => ({
          text,
          kind,
          listReference,
          alignment,
        }));

    for (const [name, properties] of [
      ["", undefined],
      [" with first-half properties", { alignment: "center" }],
    ] as const) {
      test(`reads as the merged paragraph split directly, once accepted${name}`, async () => {
        const edited = await edit("tracked-changes", properties);
        edited.acceptAll();
        expect(read(edited)).toEqual(read(await edit("direct", properties)));
      });

      test(`rejects to the document it started from, before and after a reopen${name}`, async () => {
        const fresh = read(await open());
        const edited = await edit("tracked-changes", properties);
        const reopened = await FolioDocxReviewer.fromBuffer(
          await edited.toBuffer(),
          { author: "Test", session },
        );
        for (const reviewer of [edited, reopened]) {
          reviewer.rejectAll();
          expect(read(reviewer)).toEqual(fresh);
        }
      });
    }
  },
);
