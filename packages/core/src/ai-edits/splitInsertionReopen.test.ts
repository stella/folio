/**
 * An edit inside a pending insertion leaves the insertion in stretches with
 * other content between them. The save writes each stretch as a revision of
 * its own, so the batch gives each stretch its own id when it makes it: the
 * reviewer lists, resolves, saves and reopens the same changes. On the
 * canonical engine the author's own replaced words go outright rather than
 * stay as a deletion inside the insertion, so a leading replacement leaves
 * the replacement and one stretch of the original.
 */

import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const INSERTED = "Agreement supplier buyer.";

const CASES = [
  ["leading", "Agreement", "updated", undefined],
  ["middle", "supplier", "vendor", undefined],
  ["leading", "Agreement", "updated", "Heading2"],
] as const;

type Case = (typeof CASES)[number];

const changes = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, author, text, blockId }) => ({
    type,
    author,
    text,
    blockId,
  }));

const build = async (
  session: HeadlessReviewerSession,
  [, find, replace, styleId]: Case,
): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown("# Title\n\nFirst clause.\n\nLast clause.")),
    { author: "AI", session },
  );
  const apply = (operation: Record<string, unknown>) =>
    expect(
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [{ id: String(operation["type"]), ...operation }],
        } as never,
        { undefinedReferences: "refuse" },
      ).skipped,
    ).toEqual([]);
  const first = reviewer
    .getContent()
    .find(({ text }) => text === "First clause.")?.id;
  apply({
    type: "insertAfterBlock",
    blockId: first,
    text: INSERTED,
    ...(styleId && { styleId }),
  });
  const inserted = reviewer
    .getContent()
    .find(({ text }) => text === INSERTED)?.id;
  apply({ type: "replaceInBlock", blockId: inserted, find, replace });
  return reviewer;
};

const reopen = async (
  reviewer: FolioDocxReviewer,
  session: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { session });

/**
 * Accept the paragraph's last inserted stretch, of the original insertion or
 * its replacement; every other stretch stays pending.
 */
const acceptLastStretch = (
  reviewer: FolioDocxReviewer,
  replacement: string,
) => {
  const pendingStretches = () =>
    reviewer
      .getChanges()
      .filter(
        ({ type, text }) =>
          type === "insertion" &&
          (INSERTED.includes(text) || text === replacement),
      );
  const stretches = pendingStretches();
  expect(stretches.length).toBeGreaterThan(1);
  const last = stretches.at(-1);
  if (!last) throw new Error("no insertion stretch");
  expect(reviewer.acceptChange(last)).toBe(true);
  expect(pendingStretches().map(({ text }) => text)).toEqual(
    stretches.slice(0, -1).map(({ text }) => text),
  );
  return changes(reviewer);
};

describe.each(SESSIONS)(
  "an edit inside a pending insertion (%s)",
  (session) => {
    test.each(CASES)(
      "replacing its %s words %s with %s (style %s) reads the same after a save",
      async (...entry) => {
        const reviewer = await build(session, entry);
        const before = changes(reviewer);
        expect(changes(await reopen(reviewer, session))).toEqual(before);
      },
    );

    test.each(CASES)(
      "replacing its %s words %s with %s (style %s) gives each stretch its own id",
      async (...entry) => {
        const reviewer = await build(session, entry);
        for (const current of [reviewer, await reopen(reviewer, session)]) {
          const ids = current.getChanges().map(({ id }) => id);
          expect(new Set(ids).size).toBe(ids.length);
        }
      },
    );

    test.each(CASES)(
      "replacing its %s words %s with %s (style %s), accepting one stretch leaves the others pending",
      async (...entry) => {
        const reviewer = await build(session, entry);
        const saved = await reopen(reviewer, session);
        const [, , replacement] = entry;
        const live = acceptLastStretch(reviewer, replacement);
        const reopened = acceptLastStretch(saved, replacement);
        expect(reopened).toEqual(live);
      },
    );
  },
);
