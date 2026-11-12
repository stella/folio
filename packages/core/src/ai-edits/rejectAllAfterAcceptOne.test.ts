/**
 * Rejecting every change after one of several suggestions was accepted:
 * suggestions insert a paragraph before one block, delete another and insert
 * two paragraphs before it, one change is accepted, and a last suggestion
 * inserts a right-aligned paragraph after any block.
 *
 * - The last suggestion is rejected with the rest, so it leaves nothing: the
 *   paragraphs read as they do when it was never made, each with its own
 *   style and direct alignment. Accepting a deleted paragraph break joins
 *   that paragraph's words into a suggested paragraph; rejecting both
 *   suggestions runs them into the paragraph that stays, not into the
 *   rejected one.
 * - The saved package, reopened and rejected, reads as the live reviewer
 *   rejected. Deleting the last paragraph and accepting its deleted break
 *   leaves a suggested paragraph holding another suggestion's break; the
 *   save rejects both, so the accepted join is in the package too.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const LINES = [
  "Service Agreement",
  "This agreement is made between the parties named below.",
  "The Supplier delivers the goods on time and in good order.",
  "The Buyer pays each invoice within thirty days.",
  "Signed in two copies.",
] as const;

const MARKDOWN = LINES.map((line, index) =>
  index === 0 ? `# ${line}` : line,
).join("\n\n");

const open = async (bytes: Uint8Array | ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });

const idOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((entry) => entry.text === text);
  if (!block) throw new Error(`No block reads "${text}"`);
  return block.id;
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "suggested",
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations,
    },
    { undefinedReferences: "refuse" },
  );

/** The heading, the body paragraphs and two comment threads. */
const fixture = async (): Promise<Uint8Array> => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(MARKDOWN))))
    .docx;
  const reviewer = await open(bytes);
  const commented = apply(reviewer, "direct", [
    {
      id: "c1",
      type: "commentOnBlock",
      blockId: idOf(reviewer, LINES[2]),
      comment: { text: "Define good order." },
    },
    {
      id: "c2",
      type: "commentOnBlock",
      blockId: idOf(reviewer, LINES[3]),
      comment: { text: "Thirty days is long." },
    },
  ]);
  expect(commented.applied).toHaveLength(2);
  return new Uint8Array(await reviewer.toBuffer());
};

/** What a reader sees of each paragraph, and the comments it anchors. */
const project = (reviewer: FolioDocxReviewer) => ({
  blocks: reviewer.getContent().map((block) => ({
    text: block.text,
    kind: block.kind,
    styleId: block.styleId,
    headingLevel: block.headingLevel,
    directAlignment:
      "directAlignment" in block ? block.directAlignment : undefined,
  })),
  comments: reviewer
    .getComments()
    .map((comment) => [comment.text, comment.anchoredText ?? ""]),
});

/** Every change rejected, saved and reopened. */
const rejected = async (reviewer: FolioDocxReviewer) => {
  reviewer.rejectAll();
  return project(await open(await reviewer.toBuffer()));
};

type Step = {
  /** The paragraph the second batch deletes and inserts two before. */
  deleted: number;
  /** The index of the change accepted after it. */
  accept: number;
};

/** The suggestions and the accepted change; `null` when that change does not exist. */
const suggest = async (
  bytes: Uint8Array,
  { deleted, accept }: Step,
): Promise<FolioDocxReviewer | null> => {
  const reviewer = await open(bytes);
  const first = apply(reviewer, "suggested", [
    {
      id: "before",
      type: "insertBeforeBlock",
      blockId: idOf(reviewer, LINES[1]),
      text: "Period written supplier invoice.",
    },
  ]);
  expect(first.applied).toHaveLength(1);
  const target = idOf(reviewer, LINES[deleted] ?? "");
  const second = apply(reviewer, "suggested", [
    { id: "delete", type: "deleteBlock", blockId: target },
    {
      id: "first",
      type: "insertBeforeBlock",
      blockId: target,
      text: "First added clause.",
    },
    {
      id: "second",
      type: "insertBeforeBlock",
      blockId: target,
      text: "Second added clause.",
    },
  ]);
  expect(second.applied).toHaveLength(3);
  const change = reviewer.getChanges().at(accept);
  if (change === undefined) return null;
  expect(reviewer.acceptChange(change)).toBe(true);
  return reviewer;
};

/**
 * The last suggestion: a right-aligned paragraph after the block at `after`.
 * Whether it applied: a paragraph the reviewer cannot place one after skips.
 */
const suggestAfter = (reviewer: FolioDocxReviewer, after: number): boolean => {
  const anchor = reviewer.getContent().at(after);
  if (anchor === undefined) throw new Error(`No block at ${after}`);
  const result = apply(reviewer, "suggested", [
    {
      id: "after",
      type: "insertAfterBlock",
      blockId: anchor.id,
      text: "Term agreement party invoice.",
      alignment: "right",
    },
  ]);
  return result.applied.length === 1;
};

describe("reject all after accepting one of several suggestions", () => {
  // Deleting the last paragraph leaves a suggested paragraph holding another
  // suggestion's break; a rejected paragraph suggested right before it still
  // leaves its formatting on the last paragraph (the todo below).
  test.todo(
    "a rejected paragraph suggested before one holding another suggestion's break leaves no formatting",
  );

  for (const deleted of [0, 1]) {
    test(`a rejected last suggestion leaves nothing, deleting paragraph ${deleted}`, async () => {
      const bytes = await fixture();
      let checked = 0;
      for (let accept = 0; ; accept += 1) {
        const step = { deleted, accept };
        const probe = await suggest(bytes, step);
        if (probe === null) break;
        const expected = await rejected(probe);
        const blocks = (await suggest(bytes, step))?.getContent().length ?? 0;
        for (let after = 0; after < blocks; after += 1) {
          const label = { deleted, accept, after };
          const live = await suggest(bytes, step);
          if (live === null) throw new Error("the flow no longer replays");
          if (!suggestAfter(live, after)) continue;
          const saved = await open(await live.toBuffer());
          expect({ label, ...(await rejected(live)) }).toEqual({
            label,
            ...expected,
          });
          expect({ label, ...(await rejected(saved)) }).toEqual({
            label,
            ...expected,
          });
          checked += 1;
        }
      }
      expect(checked).toBeGreaterThan(10);
    });
  }

  test("the reported flow keeps the joined words in the paragraph that stays", async () => {
    // The heading is deleted and its deleted break alone is accepted.
    const flow = await open(await fixture());
    apply(flow, "suggested", [
      {
        id: "before",
        type: "insertBeforeBlock",
        blockId: idOf(flow, LINES[1]),
        text: "Period written supplier invoice.",
      },
    ]);
    const heading = idOf(flow, LINES[0]);
    apply(flow, "suggested", [
      { id: "delete", type: "deleteBlock", blockId: heading },
      {
        id: "first",
        type: "insertBeforeBlock",
        blockId: heading,
        text: "First added clause.",
      },
      {
        id: "second",
        type: "insertBeforeBlock",
        blockId: heading,
        text: "Second added clause.",
      },
    ]);
    const deletedBreak = flow
      .getChanges()
      .find((change) => change.type === "paragraphMarkDeleted");
    if (deletedBreak === undefined) throw new Error("no deleted break");
    expect(flow.acceptChange(deletedBreak)).toBe(true);
    expect(suggestAfter(flow, 2)).toBe(true);
    const saved = await open(await flow.toBuffer());
    const joined = {
      text: `${LINES[0]}${LINES[1]}`,
      kind: "paragraph",
      styleId: undefined,
      headingLevel: undefined,
      directAlignment: undefined,
    };
    for (const reviewer of [flow, saved]) {
      const { blocks } = await rejected(reviewer);
      expect(blocks[0]).toEqual(joined);
      expect(blocks.map(({ text }) => text)).toEqual([
        `${LINES[0]}${LINES[1]}`,
        ...LINES.slice(2),
      ]);
    }
  });

  test("an accepted deleted last break is in the saved package", async () => {
    const flow = await open(await fixture());
    const last = idOf(flow, LINES[4]);
    apply(flow, "suggested", [
      { id: "delete", type: "deleteBlock", blockId: last },
      {
        id: "first",
        type: "insertBeforeBlock",
        blockId: last,
        text: "First added clause.",
      },
      {
        id: "second",
        type: "insertBeforeBlock",
        blockId: last,
        text: "Second added clause.",
      },
    ]);
    const deletedBreak = flow
      .getChanges()
      .find((change) => change.type === "paragraphMarkDeleted");
    if (deletedBreak === undefined) throw new Error("no deleted break");
    expect(flow.acceptChange(deletedBreak)).toBe(true);
    const saved = await open(await flow.toBuffer());
    const live = await rejected(flow);
    expect(await rejected(saved)).toEqual(live);
    expect(live.blocks.map(({ text }) => text)).toEqual([
      ...LINES.slice(0, 3),
      `${LINES[3]}${LINES[4]}`,
    ]);
  });
});
