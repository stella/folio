/**
 * A paragraph inserted before one whose break a tracked merge removes copies
 * the properties the joined paragraph will have, so accepting gives what the
 * same edits applied directly give.
 */

import { describe, expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN = "# Terms\n\n1. Numbered clause one\n\nUnnumbered body text.";

const open = async () => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(MARKDOWN))))
    .docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });
};

type MergedProperties = Extract<
  FolioDocumentOperation,
  { type: "mergeBlockWithNext" }
>["mergedParagraphProperties"];

const requests: Record<string, (numId: number) => MergedProperties> = {
  "the next paragraph's list": (numId) => ({
    numbering: { kind: "reference", numId, ilvl: 0 },
  }),
  "another style": () => ({ styleId: "Heading2" }),
  "no list": () => ({ numbering: { kind: "none" } }),
  "nothing more": () => undefined,
};

type EditOptions = {
  mode: "direct" | "tracked-changes";
  batched: boolean;
  merged: (numId: number) => MergedProperties;
};

/** Merge the heading into the list item, insert before it, accept, and read. */
const edit = async ({ mode, batched, merged }: EditOptions) => {
  const reviewer = await open();
  const [heading, item] = reviewer.snapshot().blocks;
  const numId = item?.listReference?.numId;
  if (!heading || numId === undefined) throw new Error("Missing fixture");
  const properties = merged(numId);
  const merge: FolioDocumentOperation = {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: heading.id,
    separator: " ",
    ...(properties !== undefined && { mergedParagraphProperties: properties }),
  };
  const insert: FolioDocumentOperation = {
    id: "insert",
    type: "insertBeforeBlock",
    blockId: heading.id,
    text: "Inserted clause.",
  };
  const batches = batched ? [[merge, insert]] : [[merge], [insert]];
  for (const operations of batches) {
    const { skipped } = reviewer.applyOperations(operations, {
      mode,
      atomic: true,
    });
    expect(skipped).toEqual([]);
  }
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { author: "Test" },
  );
  return reopened.snapshot().blocks.map((block) => ({
    text: block.text,
    styleId: block.styleId,
    statedNumbering: block.statedNumbering,
    listReference: block.listReference,
  }));
};

describe("an insertion before a paragraph a tracked merge joins", () => {
  for (const [name, merged] of Object.entries(requests)) {
    for (const batched of [true, false]) {
      test(`copies the joined paragraph's properties (${name}, ${batched ? "one batch" : "two batches"})`, async () => {
        const direct = await edit({ mode: "direct", batched, merged });
        expect(direct[0]?.text).toBe("Inserted clause.");
        expect(
          await edit({ mode: "tracked-changes", batched, merged }),
        ).toEqual(direct);
      });
    }
  }
});
