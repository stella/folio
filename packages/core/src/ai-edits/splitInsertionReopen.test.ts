/**
 * An edit inside a pending insertion leaves the insertion in stretches with
 * other content between them. The save writes each stretch as a revision of
 * its own, so the batch gives each stretch its own id when it makes it: the
 * reviewer lists, resolves, saves and reopens the same changes.
 */

import { describe, expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const INSERTED = "Agreement supplier buyer.";

const CASES = [
  ["leading", "Agreement", "updated", undefined],
  ["middle", "supplier", "vendor", undefined],
  ["leading", "Agreement", "updated", "Heading2"],
] as const;

type Case = (typeof CASES)[number];

const changes = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, author, text, blockId }) => ({ type, author, text, blockId }));

const build = async ([, find, replace, styleId]: Case): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown("# Title\n\nFirst clause.\n\nLast clause.")),
    { author: "AI" },
  );
  const apply = (operation: Record<string, unknown>) =>
    expect(
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: String(operation["type"]), ...operation }],
      } as never).skipped,
    ).toEqual([]);
  const first = reviewer.getContent().find(({ text }) => text === "First clause.")?.id;
  apply({ type: "insertAfterBlock", blockId: first, text: INSERTED, ...(styleId && { styleId }) });
  const inserted = reviewer.getContent().find(({ text }) => text === INSERTED)?.id;
  apply({ type: "replaceInBlock", blockId: inserted, find, replace });
  return reviewer;
};

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());

/** Accept the insertion's last stretch; the rest must stay pending. */
const acceptLastStretch = (reviewer: FolioDocxReviewer) => {
  const insertions = reviewer
    .getChanges()
    .filter(({ type, text }) => type === "insertion" && INSERTED.includes(text));
  expect(insertions.length).toBeGreaterThan(1);
  const last = insertions.at(-1);
  if (!last) throw new Error("no insertion stretch");
  expect(reviewer.acceptChange(last)).toBe(true);
  return changes(reviewer);
};

describe("an edit inside a pending insertion", () => {
  test.each(CASES)(
    "replacing its %s words %s with %s (style %s) reads the same after a save",
    async (...entry) => {
      const reviewer = await build(entry);
      const before = changes(reviewer);
      expect(changes(await reopen(reviewer))).toEqual(before);
    },
  );

  test.each(CASES)(
    "replacing its %s words %s with %s (style %s) gives each stretch its own id",
    async (...entry) => {
      const reviewer = await build(entry);
      for (const current of [reviewer, await reopen(reviewer)]) {
        const ids = current.getChanges().map(({ id }) => id);
        expect(new Set(ids).size).toBe(ids.length);
      }
    },
  );

  test.each(CASES)(
    "replacing its %s words %s with %s (style %s), accepting one stretch leaves the others pending",
    async (...entry) => {
      const reviewer = await build(entry);
      const saved = await reopen(reviewer);
      const live = acceptLastStretch(reviewer);
      const reopened = acceptLastStretch(saved);
      expect(live.some(({ type, text }) => type === "insertion" && INSERTED.startsWith(text))).toBe(
        true,
      );
      expect(reopened).toEqual(live);
    },
  );
});
