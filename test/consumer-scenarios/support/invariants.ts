/**
 * The invariants every scenario checks after it acts: the package saves, it
 * reopens to what the reviewer showed before the save, and every reader of
 * the saved package agrees on its blocks.
 */

import assert from "node:assert/strict";

import { openReviewer } from "./documents.ts";
import { TOLERATED_DISAGREEMENTS } from "./known-issues.ts";
import { type BlockView, labelFields, readAll } from "./readers.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

export type VisibleStateOptions = {
  /**
   * Compare the fields a reviewer only settles at open too: list labels and
   * levels (STALE_LIST_LABELS) and comment anchor text (COMMENT_ANCHOR_DRIFT
   * in known-issues.ts). Off by default; the saved package's readers are
   * compared with each other either way.
   */
  exact?: boolean;
};

/** Everything a person or a model can see of a reviewer's document. */
export const visibleState = (reviewer: Reviewer, { exact = false }: VisibleStateOptions = {}) => ({
  blocks: reviewer.getContent().map((block) =>
    Object.assign(
      {
        id: block.id,
        // #1094: a numbered heading reads as a list item once its numbering
        // is resolved, which a reviewer only does at open.
        kind:
          !exact && block.kind === "listItem" && block.headingLevel !== undefined
            ? "heading"
            : block.kind,
        text: block.text,
        headingLevel: block.headingLevel,
      },
      exact ? { displayLabel: block.displayLabel, listLevel: block.listLevel } : {},
    ),
  ),
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
    anchor: exact ? comment.anchoredText : comment.anchoredText.length > 0,
    done: comment.done,
    replies: comment.replies.map((reply) => reply.text),
  })),
  notes: reviewer.getNotesAsText(),
});

const describeError = (error: unknown): string =>
  error instanceof Error ? error.message.split("\n").slice(0, 3).join(" | ") : String(error);

export type SaveOptions = VisibleStateOptions & {
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
  assert.deepEqual(
    visibleState(reopened, options),
    options.persisted ?? visibleState(reviewer, options),
    `${context}: the reopened package shows something else than the reviewer that saved it`,
  );
  return { bytes, reopened };
};

/**
 * #1094: `getContent()` and the snapshot classify by numbering before
 * heading: a numbered heading reads as `listItem` while Markdown shows a
 * heading, and a paragraph whose numbering shows no marker (a level the
 * instance does not define) reads as `listItem` while Markdown shows plain
 * text. Tolerated, and pinned by an expected failure, until the fix lands.
 */
const tolerateNumberedHeadingKind = (block: BlockView): BlockView => {
  if (block.kind !== "listItem") return block;
  if (block.headingLevel !== undefined) return { ...block, kind: "heading" };
  return block.number === undefined ? { ...block, kind: "paragraph" } : block;
};

export type ReaderAgreementOptions = {
  /** Fail on disagreements an open issue already reports. */
  strict?: boolean;
};

/** Every reader of `bytes` shows the same blocks, kinds, levels and numbers. */
export const assertReadersAgree = async (
  bytes: Uint8Array,
  context: string,
  { strict = false }: ReaderAgreementOptions = {},
): Promise<void> => {
  const views = await readAll(bytes);
  const tolerate =
    !strict && TOLERATED_DISAGREEMENTS.has(1094)
      ? (blocks: BlockView[]) => blocks.map(tolerateNumberedHeadingKind)
      : (blocks: BlockView[]) => blocks;

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
  if (strict) {
    // A model is given the numbers a reader sees (#1094).
    assert.deepEqual(
      views.rows.map(labelFields),
      views.labels,
      `${context}: read_document labels vs getContent()`,
    );
  }
  assert.deepEqual(
    views.markdown,
    tolerate(views.getContentAsMarkdown),
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
