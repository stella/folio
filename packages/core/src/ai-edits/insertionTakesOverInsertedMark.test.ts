import { panic } from "better-result";
import { expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

// A body of one paragraph S. A tracked table after S, then a tracked
// paragraph P after S: P takes over S's pending inserted mark. Each
// operation reviews on its own: accepting some and rejecting the others, in
// either order, reads as the accepted ones made directly.

const SOURCE = "Source paragraph.";
const ADDED = "Inserted clause.";

type Step = (source: string) => FolioDocumentOperation;

const TABLE: Step = (blockId) => ({
  id: "table",
  type: "insertTable",
  blockId,
  rows: [["Term", "Value"]],
});
const PARAGRAPH: Step = (blockId) => ({
  id: "paragraph",
  type: "insertAfterBlock",
  blockId,
  text: ADDED,
});
const STEPS = [TABLE, PARAGRAPH] as const;

const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical", "legacy"];

const open = async (session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(await createDocx(fromMarkdown(SOURCE)), {
    session,
  });

const sourceOf = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().find(({ text }) => text === SOURCE)?.id ??
  panic("The fixture needs its source paragraph.");

type ApplyOptions = {
  reviewer: FolioDocxReviewer;
  mode: FolioAIEditApplyMode;
  step: Step;
};

const applyStep = ({ reviewer, mode, step }: ApplyOptions) => {
  const result = reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [step(sourceOf(reviewer))] },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return result.undoHandle ?? panic("An admitted batch can be undone.");
};

/** What a reader sees of the body: each block's text. */
const read = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

const pendingIds = (reviewer: FolioDocxReviewer) =>
  new Set(reviewer.getChanges().map(({ id }) => id));

type Tracked = {
  reviewer: FolioDocxReviewer;
  /** The revision ids each operation added, in operation order. */
  own: readonly (readonly number[])[];
};

const tracked = async (session: HeadlessReviewerSession): Promise<Tracked> => {
  const reviewer = await open(session);
  const own: number[][] = [];
  for (const step of STEPS) {
    const before = pendingIds(reviewer);
    applyStep({ reviewer, mode: "tracked-changes", step });
    own.push([...pendingIds(reviewer)].filter((id) => !before.has(id)));
  }
  return { reviewer, own };
};

const directOf = async (
  session: HeadlessReviewerSession,
  accepted: readonly boolean[],
) => {
  const reviewer = await open(session);
  for (const [index, step] of STEPS.entries())
    if (accepted[index] === true) applyStep({ reviewer, mode: "direct", step });
  return read(reviewer);
};

type ResolveOptions = {
  reviewer: FolioDocxReviewer;
  ids: readonly number[];
  accept: boolean;
};

const MARK_CHANGES: ReadonlySet<string> = new Set([
  "paragraphMarkInserted",
  "paragraphMarkDeleted",
]);

/**
 * Resolves the operation's changes one at a time, as a reviewer does: words
 * and rows before paragraph marks, and a change that waits on another (a
 * mark before a pending table) again once the others are resolved.
 */
const resolve = ({ reviewer, ids, accept }: ResolveOptions) => {
  for (;;) {
    const pending = reviewer
      .getChanges()
      .filter(({ id }) => ids.includes(id))
      .toSorted(
        (left, right) =>
          Number(MARK_CHANGES.has(left.type)) -
          Number(MARK_CHANGES.has(right.type)),
      )
      .map(({ id }) => id);
    const resolved = pending.filter((id) =>
      accept ? reviewer.acceptChange(id) : reviewer.rejectChange(id),
    );
    if (resolved.length === 0) return;
  }
};

const DECISIONS = [
  [true, true],
  [true, false],
  [false, true],
  [false, false],
] as const;

for (const session of SESSIONS) {
  test(`${session}: the paragraph after a paragraph whose table ends the body stands before the table`, async () => {
    const { reviewer, own } = await tracked(session);
    expect(read(reviewer).slice(0, 4)).toEqual([
      SOURCE,
      ADDED,
      "Term",
      "Value",
    ]);
    // The paragraph operation added its own records.
    expect(own[1]?.length ?? 0).toBeGreaterThan(0);
  });

  for (const accepted of DECISIONS)
    for (const order of ["in order", "in reverse"] as const)
      test(`${session}: accepting [${accepted.join(", ")}] of the table and the paragraph ${order} reads as the accepted ones made directly`, async () => {
        const { reviewer, own } = await tracked(session);
        const indices = order === "in order" ? [0, 1] : [1, 0];
        for (const index of indices)
          resolve({
            reviewer,
            ids: own[index] ?? [],
            accept: accepted[index] === true,
          });
        expect(reviewer.getChanges()).toEqual([]);
        const expected = await directOf(session, accepted);
        expect(read(reviewer)).toEqual(expected);
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        expect(read(reopened)).toEqual(expected);
      });

  test(`${session}: rejecting the table alone keeps the paragraph's id`, async () => {
    const { reviewer, own } = await tracked(session);
    const added =
      reviewer.getContent().find(({ text }) => text === ADDED)?.id ??
      panic("The paragraph was added.");
    resolve({ reviewer, ids: own[0] ?? [], accept: false });
    expect(reviewer.getContent().find(({ text }) => text === ADDED)?.id).toBe(
      added,
    );
  });

  test(`${session}: undoing each operation gives back the state before it`, async () => {
    const reviewer = await open(session);
    const original = await reviewer.toBuffer();
    const originalRead = read(reviewer);
    const tableUndo = applyStep({
      reviewer,
      mode: "tracked-changes",
      step: TABLE,
    });
    const afterTable = read(reviewer);
    const afterTableChanges = reviewer.getChanges();
    const paragraphUndo = applyStep({
      reviewer,
      mode: "tracked-changes",
      step: PARAGRAPH,
    });
    expect(reviewer.undoDocumentOperations(paragraphUndo).status).toBe(
      "undone",
    );
    expect(read(reviewer)).toEqual(afterTable);
    expect(reviewer.getChanges()).toEqual(afterTableChanges);
    expect(reviewer.undoDocumentOperations(tableUndo).status).toBe("undone");
    expect(read(reviewer)).toEqual(originalRead);
    expect(reviewer.getChanges()).toEqual([]);
    expect(
      read(await FolioDocxReviewer.fromBuffer(original, { session })),
    ).toEqual(originalRead);
  });
}
