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
import { FolioDocxReviewer, type FolioReviewChangeKind } from "./headless";

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

type ApplyOptions = {
  reviewer: FolioDocxReviewer;
  mode: "direct" | "suggested";
  operations: FolioDocumentOperation[];
};

const apply = ({ reviewer, mode, operations }: ApplyOptions) =>
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
  const commented = apply({
    reviewer,
    mode: "direct",
    operations: [
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
    ],
  });
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

/**
 * The suggestions and the accepted change, with the kind of change accepted;
 * `null` when that change does not exist.
 */
const suggest = async (
  bytes: Uint8Array,
  { deleted, accept }: Step,
): Promise<{
  reviewer: FolioDocxReviewer;
  accepted: FolioReviewChangeKind;
} | null> => {
  const reviewer = await open(bytes);
  const first = apply({
    reviewer,
    mode: "suggested",
    operations: [
      {
        id: "before",
        type: "insertBeforeBlock",
        blockId: idOf(reviewer, LINES[1]),
        text: "Period written supplier invoice.",
      },
    ],
  });
  expect(first.applied).toHaveLength(1);
  const target = idOf(reviewer, LINES[deleted] ?? "");
  const second = apply({
    reviewer,
    mode: "suggested",
    operations: [
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
    ],
  });
  expect(second.applied).toHaveLength(3);
  const change = reviewer.getChanges().at(accept);
  if (change === undefined) return null;
  expect(reviewer.acceptChange(change)).toBe(true);
  return { reviewer, accepted: change.type };
};

/**
 * The last suggestion: a right-aligned paragraph after the block at `after`.
 * Whether it applied: a paragraph the reviewer cannot place one after skips.
 */
const suggestAfter = (reviewer: FolioDocxReviewer, after: number): boolean => {
  const anchor = reviewer.getContent().at(after);
  if (anchor === undefined) throw new Error(`No block at ${after}`);
  const result = apply({
    reviewer,
    mode: "suggested",
    operations: [
      {
        id: "after",
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Term agreement party invoice.",
        alignment: "right",
      },
    ],
  });
  return result.applied.length === 1;
};

type Shape = Step & {
  /** The index of the block the last suggestion inserts after. */
  after: number;
};

/**
 * KNOWN GAP, held by the `test.failing` below: the last paragraph is deleted,
 * its deleted break alone is accepted, and the last suggestion inserts a
 * paragraph right before the suggested paragraph that then holds another
 * suggestion's displaced break (the first added clause). Rejecting every
 * change still leaves that suggestion's alignment and a break on the last
 * original paragraph, live and in the saved package.
 */
type DisplacedBreakGapOptions = {
  shape: Shape;
  accepted: FolioReviewChangeKind;
  reviewer: FolioDocxReviewer;
};

const isDisplacedBreakGap = ({
  shape: { deleted, after },
  accepted,
  reviewer,
}: DisplacedBreakGapOptions): boolean =>
  deleted === LINES.length - 1 &&
  accepted === "paragraphMarkDeleted" &&
  reviewer.getContent().at(after + 1)?.text === "First added clause.";

/**
 * Whether rejecting every change after the last suggestion reads, live and
 * saved, as rejecting every change without it; `null` when the shape does not
 * apply (no such change, or no paragraph can follow that block).
 */
const rejectsToBaseline = async (
  bytes: Uint8Array,
  shape: Shape,
): Promise<{ matches: boolean; gap: boolean } | null> => {
  const probe = await suggest(bytes, shape);
  if (probe === null) return null;
  const expected = await rejected(probe.reviewer);
  const live = await suggest(bytes, shape);
  if (live === null) throw new Error("the flow no longer replays");
  const gap = isDisplacedBreakGap({
    shape,
    accepted: live.accepted,
    reviewer: live.reviewer,
  });
  if (!suggestAfter(live.reviewer, shape.after)) return null;
  const saved = await open(await live.reviewer.toBuffer());
  const results = [await rejected(live.reviewer), await rejected(saved)].map(
    (result) => JSON.stringify(result),
  );
  return {
    matches: results.every((result) => result === JSON.stringify(expected)),
    gap,
  };
};

/** Every shape over the paragraph `deleted`: each accepted change, each anchor. */
const shapesOf = async (
  bytes: Uint8Array,
  deleted: number,
): Promise<Shape[]> => {
  const shapes: Shape[] = [];
  for (let accept = 0; ; accept += 1) {
    const probe = await suggest(bytes, { deleted, accept });
    if (probe === null) return shapes;
    const blocks = probe.reviewer.getContent().length;
    for (let after = 0; after < blocks; after += 1) {
      shapes.push({ deleted, accept, after });
    }
  }
};

describe("reject all after accepting one of several suggestions", () => {
  for (const deleted of LINES.keys()) {
    test(`a rejected last suggestion leaves nothing, deleting paragraph ${deleted}`, async () => {
      const bytes = await fixture();
      const mismatches: Shape[] = [];
      let checked = 0;
      let gaps = 0;
      for (const shape of await shapesOf(bytes, deleted)) {
        const outcome = await rejectsToBaseline(bytes, shape);
        if (outcome === null) continue;
        if (outcome.gap) {
          gaps += 1;
          continue;
        }
        if (!outcome.matches) mismatches.push(shape);
        checked += 1;
      }
      expect(mismatches).toEqual([]);
      expect(checked).toBeGreaterThan(10);
      // The known gap is excluded only where its predicate names it.
      expect(gaps > 0).toBe(deleted === LINES.length - 1);
    });
  }

  // KNOWN GAP (see `isDisplacedBreakGap`): passes while the gap is there and
  // fails once it is fixed, so the predicate and this test go together.
  test.failing(
    "a rejected paragraph suggested before one holding another suggestion's displaced break leaves nothing",
    async () => {
      const bytes = await fixture();
      const mismatches: Shape[] = [];
      for (const shape of await shapesOf(bytes, LINES.length - 1)) {
        const outcome = await rejectsToBaseline(bytes, shape);
        if (outcome?.gap === true && !outcome.matches) mismatches.push(shape);
      }
      expect(mismatches).toEqual([]);
    },
  );

  test("the reported flow keeps the joined words in the paragraph that stays", async () => {
    // The heading is deleted and its deleted break alone is accepted.
    const flow = await open(await fixture());
    apply({
      reviewer: flow,
      mode: "suggested",
      operations: [
        {
          id: "before",
          type: "insertBeforeBlock",
          blockId: idOf(flow, LINES[1]),
          text: "Period written supplier invoice.",
        },
      ],
    });
    const heading = idOf(flow, LINES[0]);
    apply({
      reviewer: flow,
      mode: "suggested",
      operations: [
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
      ],
    });
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
    apply({
      reviewer: flow,
      mode: "suggested",
      operations: [
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
      ],
    });
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
