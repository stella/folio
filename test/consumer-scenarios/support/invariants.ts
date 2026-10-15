/**
 * The invariants every scenario checks after it acts: the package saves, it
 * reopens to what the reviewer showed before the save, and every reader of
 * the saved package agrees on its blocks.
 */

import assert from "node:assert/strict";
import { validateDocxPackage } from "@stll/docx-core";

import { openReviewer } from "./documents.ts";
import { projectContentPair } from "./identity.ts";
import { labelFields, readAll } from "./readers.ts";
import { assertReviewRoundTrip } from "./review-comparison.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

/** Match full story text from the end so label-like text inside a story stays literal. */
export const assertStoryInventory = (
  reviewer: Pick<Reviewer, "listStories" | "getNotesAsText">,
): void => {
  const message = "story trailer and live story inventory disagree";
  let trailer = reviewer.getNotesAsText();
  for (const { handle, text } of reviewer.listStories().toReversed()) {
    if (handle.type === "main" || text.length === 0) continue;
    const suffix = `] ${text}`;
    assert.ok(trailer.endsWith(suffix), message);
    const labelEnd = trailer.length - suffix.length + 1;
    const labelStart = trailer.lastIndexOf("\n[", labelEnd - 1) + 1;
    const label = trailer.slice(labelStart, labelEnd);
    switch (handle.type) {
      case "footnote":
      case "endnote":
        assert.equal(label, `[${handle.type} #${handle.noteId}]`, message);
        break;
      case "header":
      case "footer":
        assert.ok(label.startsWith(`[${handle.type} `), message);
        break;
      default: {
        const unreachable: never = handle;
        return unreachable;
      }
    }
    trailer = trailer.slice(0, labelStart).replace(/\n$/u, "");
  }
  assert.equal(trailer, "", message);
};

/** Everything a person or a model can see of a reviewer's document. */
export const visibleState = (reviewer: Reviewer) => ({
  // The number or bullet beside a block is part of what a reader sees: after
  // an operation adds, removes or renumbers list items, the reviewer shows
  // the numbers the saved package opens with.
  blocks: reviewer.getContent().map((block) => {
    switch (block.kind) {
      case "paragraph":
      case "heading":
      case "listItem":
        if (block.listReference === undefined) {
          return {
            id: block.id,
            kind: block.kind,
            text: block.text,
            idStability: block.idStability,
            headingLevel: block.headingLevel,
            displayLabel: block.displayLabel,
            statedNumbering: block.statedNumbering,
          };
        }
        return {
          id: block.id,
          kind: block.kind,
          text: block.text,
          idStability: block.idStability,
          headingLevel: block.headingLevel,
          displayLabel: block.displayLabel,
          statedNumbering: block.statedNumbering,
          listReference: block.listReference,
        };
      case "diagnostic":
        return {
          id: block.id,
          kind: block.kind,
          text: block.text,
          idStability: block.idStability,
          diagnostic: block.diagnostic,
        };
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }),
  // Overview only; the shared review comparator checks lossless coverage,
  // authorship dates and both complete resolution documents below.
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
  const validity = await validateDocxPackage(buffer);
  assert.equal(
    validity.valid,
    true,
    `${context}: saved OOXML invalid: ${validity.valid ? "" : validity.error}`,
  );
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
  if (options.persisted === undefined)
    await assertReviewRoundTrip({ live: reviewer, reopened, context });
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
