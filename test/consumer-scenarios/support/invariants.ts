/**
 * The invariants every scenario checks after it acts: the package saves, it
 * reopens to what the reviewer showed before the save, and every reader of
 * the saved package agrees on its blocks.
 */

import assert from "node:assert/strict";

import { openReviewer } from "./documents.ts";
import { projectContentPair } from "./identity.ts";
import { labelFields, readAll } from "./readers.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

/** Everything a person or a model can see of a reviewer's document. */
export const visibleState = (reviewer: Reviewer) => ({
  // The number or bullet beside a block is part of what a reader sees: after
  // an operation adds, removes or renumbers list items, the reviewer shows
  // the numbers the saved package opens with.
  blocks: reviewer.getContent().map((block) => ({
    id: block.id,
    kind: block.kind,
    text: block.text,
    idStability: block.idStability,
    headingLevel: block.headingLevel,
    displayLabel: block.displayLabel,
    listLevel: block.listLevel,
  })),
  // Which kinds of change by whom. How a reader groups revisions into
  // entries (a nested `w:ins > w:del` reads as one entry after a reopen, two
  // before) is not what a save must keep; what they resolve to is checked by
  // accept-all / reject-all.
  changes: [
    ...new Set(reviewer.getChanges().map((change) => `${change.type} by ${change.author}`)),
  ].sort(),
  comments: reviewer.getComments().map((comment) => ({
    author: comment.author,
    text: comment.text,
    anchor: comment.anchoredText,
    done: comment.done,
    replies: comment.replies.map((reply) => reply.text),
  })),
  notes: reviewer.getNotesAsText(),
});

const describeError = (error: unknown): string =>
  error instanceof Error ? error.message.split("\n").slice(0, 3).join(" | ") : String(error);

export type SaveOptions = {
  /**
   * What the saved package must show when it is not what the reviewer shows:
   * `"suggested"` edits stay out of the package until accepted, so a save
   * persists the state before them.
   */
  persisted?: ReturnType<typeof visibleState>;
  /** Skip the comparison (the package must still save and reopen). */
  compare?: boolean;
};

/** `toBuffer()`, reopen, and compare what the two reviewers show. */
export const saveAndReopen = async (
  reviewer: Reviewer,
  context: string,
  options: SaveOptions = {},
): Promise<{ bytes: Uint8Array; reopened: Reviewer }> => {
  let buffer: ArrayBuffer;
  try {
    buffer = await reviewer.toBuffer();
  } catch (error) {
    throw new Error(`${context}: toBuffer() threw: ${describeError(error)}`, { cause: error });
  }
  const bytes = new Uint8Array(buffer);
  let reopened: Reviewer;
  try {
    reopened = await openReviewer(bytes);
  } catch (error) {
    throw new Error(`${context}: the saved package does not reopen: ${describeError(error)}`, {
      cause: error,
    });
  }
  if (options.compare === false) {
    return { bytes, reopened };
  }
  const expected = options.persisted ?? visibleState(reviewer);
  const actual = visibleState(reopened);
  const liveStableIds = new Set(
    reviewer
      .getContent()
      .filter(({ idStability }) => idStability !== "positional")
      .map(({ id }) => id),
  );
  const blocks = projectContentPair({
    leftRows: expected.blocks,
    rightRows: actual.blocks,
    stableIds: liveStableIds,
  });
  assert.deepEqual(
    { ...actual, blocks: blocks.right },
    { ...expected, blocks: blocks.left },
    `${context}: the reopened package shows something else than the reviewer that saved it`,
  );
  return { bytes, reopened };
};

/** Every reader of `bytes` shows the same blocks, kinds, levels and numbers. */
export const assertReadersAgree = async (bytes: Uint8Array, context: string): Promise<void> => {
  const views = await readAll(bytes);

  assert.deepEqual(views.snapshot, views.getContent, `${context}: snapshot vs getContent()`);
  assert.deepEqual(
    views.rows.map((row) => row["blockId"]),
    views.ids,
    `${context}: read_document block ids vs getContent()`,
  );
  assert.deepEqual(
    views.readDocument,
    views.getContent.map(({ text, kind }) => ({ text, kind })),
    `${context}: read_document rows vs getContent()`,
  );
  // A model is given the numbers a reader sees.
  assert.deepEqual(
    views.rows.map(labelFields),
    views.labels,
    `${context}: read_document labels vs getContent()`,
  );
  assert.deepEqual(
    views.markdown,
    views.getContentAsMarkdown,
    `${context}: docxToMarkdown vs getContent()`,
  );
};

/** Save, reopen, and check the readers of the saved package. */
export const assertHealthy = async (
  reviewer: Reviewer,
  context: string,
  options: SaveOptions = {},
): Promise<{ bytes: Uint8Array; reopened: Reviewer }> => {
  const saved = await saveAndReopen(reviewer, context, options);
  await assertReadersAgree(saved.bytes, context);
  return saved;
};
