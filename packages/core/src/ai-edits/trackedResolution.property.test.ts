/**
 * Resolving tracked work is the other half of making it: whatever sequence of
 * tracked or suggested operations a reviewer applies, over pending revisions
 * of its own, resolving it must land somewhere well defined.
 *
 * For a random sequence of block operations (insert, delete, split, merge,
 * replace, table insert/delete, rows) and comments, applied one batch at a
 * time in `tracked-changes` or `suggested` mode:
 *
 * - nothing throws: applying, `rejectAll`, `acceptAll`, saving;
 * - `rejectAll` gives back the original document;
 * - `acceptAll` leaves suggested operations pending and out of a save;
 * - tracked `acceptAll` gives what the same operations give in `direct` mode,
 *   when each has a direct counterpart (its target exists there too);
 * - the resolved tracked document survives a save;
 * - a pending (tracked) document saved and reopened resolves to the same
 *   result both ways as the live one;
 * - a comment's anchored text reads the same before and after a save.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  isFolioDocumentOperationModeSupported,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph, table } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";

setDefaultTimeout(propertyTestTimeout(120_000));

const MODES = ["tracked-changes", "suggested"] as const;
type Mode = (typeof MODES)[number] | "direct";

const KINDS = [
  "insertAfterBlock",
  "insertBeforeBlock",
  "deleteBlock",
  "splitBlock",
  "mergeBlockWithNext",
  "replaceBlock",
  "replaceInBlock",
  "insertTable",
  "deleteTable",
  "insertTableRow",
  "deleteTableRow",
  "commentOnRange",
  "commentOnBlock",
] as const;
type Kind = (typeof KINDS)[number];

const SENTENCES = [
  "Inserted clause.",
  "Payment is due.",
  "New term here.",
  "Short.",
] as const;

let base: ArrayBuffer | null = null;
const baseDocument = async (): Promise<ArrayBuffer> => {
  if (base) {
    return base;
  }
  const document = fromMarkdown(
    [
      "# Service Agreement",
      "This agreement is made between the parties named below.",
      "The Supplier delivers the goods on time and in good order.",
      "The Buyer pays each invoice within thirty days.",
      "Signed in two copies.",
    ].join("\n\n"),
  );
  document.package.document.content.splice(
    3,
    0,
    table({ header: ["Item", "Price"], rows: [["Widget", "10"]] }),
    paragraph("Prices exclude taxes."),
  );
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  base = docx.buffer.slice(
    docx.byteOffset,
    docx.byteOffset + docx.byteLength,
  ) as ArrayBuffer;
  return base;
};

const open = async (
  bytes?: ArrayBuffer,
  session?: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(bytes ?? (await baseDocument()), {
    author: "Reviewer",
    ...(session !== undefined && { session }),
  });

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

type Block = ReturnType<FolioDocxReviewer["getContent"]>[number];

/** What a reader is left with: every block's kind and text. */
const read = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map((block) => `${block.kind}: ${block.text}`);

const anchors = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getComments().map((comment) => comment.anchoredText);

const reopen = async (
  reviewer: FolioDocxReviewer,
  session?: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> => open(await reviewer.toBuffer(), session);

const words = (text: string) =>
  [...text.matchAll(/[A-Za-z]{3,}/gu)].map((match) => ({
    word: match[0],
    start: match.index,
  }));

/**
 * `follow` aims at the neighbourhood of the previous step (`pick` says which
 * block): its target, the blocks it added, the paragraphs either side, and the
 * story's last paragraph. Steps then compose where they interfere: a split
 * and a merge of its half, an insertion and its deletion, a merge into a
 * deleted or inserted neighbour, a comment and a replacement of its words.
 */
type Intent = { kind: Kind; pick: number; detail: number; follow: boolean };

/** A concrete operation for `intent` against the blocks the reviewer shows now. */
const operationFor = (
  blocks: readonly Block[],
  { kind, pick, detail, follow }: Intent,
  focus: readonly string[],
) => {
  const body = blocks.filter((block) => block.table === undefined);
  const cells = blocks.filter((block) => block.table !== undefined);
  const worded = body.filter((block) => words(block.text).length > 0);
  const choose = (from: readonly Block[]): Block | undefined =>
    (follow
      ? from.find(({ id }) => id === focus[pick % Math.max(focus.length, 1)])
      : undefined) ?? from[pick % Math.max(from.length, 1)];
  const sentence = SENTENCES[detail % SENTENCES.length];
  switch (kind) {
    case "insertAfterBlock":
    case "insertBeforeBlock": {
      const block = choose(body);
      return block && { type: kind, blockId: block.id, text: sentence };
    }
    case "deleteBlock": {
      const block = choose(body);
      return body.length > 2 && block
        ? { type: kind, blockId: block.id }
        : undefined;
    }
    case "replaceBlock": {
      const block = choose(body);
      return block && { type: kind, blockId: block.id, text: sentence };
    }
    case "replaceInBlock": {
      const block = choose(worded);
      const found =
        block && words(block.text)[detail % words(block.text).length];
      return (
        block &&
        found && {
          type: kind,
          blockId: block.id,
          find: found.word,
          replace: sentence,
        }
      );
    }
    case "splitBlock": {
      const block = choose(
        worded.filter((candidate) => words(candidate.text).length > 1),
      );
      const at =
        block &&
        words(block.text).slice(1)[detail % (words(block.text).length - 1)];
      return block && at && { type: kind, blockId: block.id, offset: at.start };
    }
    case "mergeBlockWithNext": {
      const candidates = body.filter((block) => {
        const next = blocks[blocks.indexOf(block) + 1];
        return next !== undefined && next.table === undefined;
      });
      const block = choose(candidates);
      return block && { type: kind, blockId: block.id, separator: " " };
    }
    case "insertTable": {
      const block = choose(body);
      return (
        block && { type: kind, blockId: block.id, rows: [["Term", "Value"]] }
      );
    }
    case "deleteTable":
    case "deleteTableRow": {
      const cell = choose(cells);
      return cell && { type: kind, blockId: cell.id };
    }
    case "insertTableRow": {
      const cell = choose(cells);
      return cell && { type: kind, blockId: cell.id, position: "after" };
    }
    case "commentOnBlock": {
      // A quoted word, or the whole paragraph.
      const block = choose(worded);
      const quote =
        block && words(block.text)[detail % words(block.text).length];
      return (
        block &&
        quote && {
          type: kind,
          blockId: block.id,
          ...(detail % 2 === 0 && { quote: quote.word }),
          comment: { text: "Why?" },
        }
      );
    }
    case "commentOnRange": {
      const block = choose(worded);
      const found =
        block && words(block.text)[detail % words(block.text).length];
      const range =
        block &&
        found &&
        createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: found.start,
          endOffset: found.start + found.word.length,
        });
      return range
        ? { type: kind, range, comment: { text: "Why?" } }
        : undefined;
    }
  }
};

type Operation = NonNullable<ReturnType<typeof operationFor>>;

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operation: Operation,
): boolean =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "1", ...operation }],
    } as never,
    { undefinedReferences: "refuse" },
  ).applied.length === 1;

const idsOf = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map((block) => block.id);

/**
 * An applied operation, the text its target read, whether a table followed
 * the target, and the blocks it added.
 */
type Step = {
  operation: Operation;
  targetText: string;
  beforeTable: boolean;
  added: string[];
};

const targetOf = (operation: Operation): string =>
  "blockId" in operation ? operation.blockId : operation.range.blockId;

const initialFocus = (reviewer: FolioDocxReviewer): string[] =>
  reviewer
    .getContent()
    .slice(-1)
    .map(({ id }) => id);

/** Apply `intent` when the mode supports it: the step it made, and the focus next. */
const stepOnce = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  intent: Intent,
  focus: string[],
): { step?: Step; focus: string[] } => {
  if (!isFolioDocumentOperationModeSupported(intent.kind, mode)) {
    return { focus };
  }
  const before = new Set(idsOf(reviewer));
  const blocks = reviewer.getContent();
  const operation = operationFor(blocks, intent, focus);
  const index = operation
    ? blocks.findIndex(({ id }) => id === targetOf(operation))
    : -1;
  const targetText = blocks[index]?.text;
  const beforeTable = blocks[index + 1]?.table !== undefined;
  if (
    !operation ||
    targetText === undefined ||
    !apply(reviewer, mode, operation)
  ) {
    return { focus };
  }
  const added = idsOf(reviewer).filter((id) => !before.has(id));
  const after = reviewer.getContent();
  const at = after.findIndex(({ id }) => id === targetOf(operation));
  return {
    step: { operation, targetText, beforeTable, added },
    focus: [
      targetOf(operation),
      ...added,
      ...[after[at - 1], after[at + 1], after.at(-1)].flatMap((block) =>
        block ? [block.id] : [],
      ),
    ],
  };
};

/** Apply each intent the mode supports; what each applied step added. */
const run = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  intents: readonly Intent[],
): Step[] => {
  const steps: Step[] = [];
  let focus = initialFocus(reviewer);
  for (const intent of intents) {
    const done = stepOnce(reviewer, mode, intent, focus);
    focus = done.focus;
    if (done.step) steps.push(done.step);
  }
  return steps;
};

/**
 * `run` in tracked-changes mode on a fresh document, saved and reopened
 * before each intent `saves` marks (and at the end, past the last intent)
 * when `reopens`. A reopened document mints fresh ids from what it holds,
 * not from ids the live session retired, so `minted` names each block a step
 * adds after the first such point by the step and its order there; every id
 * the first save held stays as it was.
 */
const runSaving = async (
  intents: readonly Intent[],
  saves: readonly boolean[],
  reopens: boolean,
): Promise<{ reviewer: FolioDocxReviewer; minted: Map<string, string> }> => {
  let reviewer = await open();
  let focus = initialFocus(reviewer);
  let saved = false;
  const minted = new Map<string, string>();
  for (const [index, intent] of [...intents, undefined].entries()) {
    if (saves[index] === true) {
      saved = true;
      if (reopens) reviewer = await reopen(reviewer);
    }
    if (!intent) continue;
    const done = stepOnce(reviewer, "tracked-changes", intent, focus);
    focus = done.focus;
    if (saved)
      done.step?.added.forEach((id, order) =>
        minted.set(id, `minted by step ${index} #${order}`),
      );
  }
  return { reviewer, minted };
};

/** What a reader is left with, paragraph ids included: each block, and where each comment sits. */
const identities = (
  reviewer: FolioDocxReviewer,
  aliases: ReadonlyMap<string, string> = new Map(),
): string[] => {
  const name = (id: string | null) => (id && aliases.get(id)) ?? id;
  return [
    ...reviewer
      .getContent()
      .map((block) => `${name(block.id)} ${block.kind}: ${block.text}`),
    ...reviewer
      .getComments()
      .map(
        (comment) =>
          `comment on ${name(comment.blockId)}: ${comment.anchoredText}`,
      ),
  ];
};

/**
 * The same steps in `direct` mode, or `null` when one has no counterpart
 * there: its target is gone (a block only a pending deletion still shows), or
 * direct mode refuses it. New block ids are random, so the blocks each step
 * adds are paired in order.
 */
const runDirectly = async (
  steps: readonly Step[],
): Promise<FolioDocxReviewer | null> => {
  const direct = await open();
  const ids = new Map(idsOf(direct).map((id) => [id, id]));
  for (const { operation, targetText, beforeTable, added } of steps) {
    const mapped = ids.get(targetOf(operation));
    const block = direct.getContent().find(({ id }) => id === mapped);
    if (!mapped || !block || block.text !== targetText) {
      // Gone, or not the same block any more: a pending merge still shows
      // two paragraphs where the direct run has one.
      return null;
    }
    if (operation.type === "deleteBlock" && beforeTable) {
      // A paragraph before a table whose break is itself a pending insertion
      // cannot carry an inserted and deleted mark at once, and there is no
      // paragraph to join.
      return null;
    }
    let counterpart: Operation;
    if ("range" in operation) {
      const range = createFolioAITextRangeHandle({
        blockId: mapped,
        text: block.text,
        startOffset: operation.range.startOffset,
        endOffset: operation.range.endOffset,
      });
      if (!range) {
        return null;
      }
      counterpart = { ...operation, range };
    } else {
      counterpart = { ...operation, blockId: mapped };
    }
    const before = new Set(idsOf(direct));
    if (!apply(direct, "direct", counterpart)) {
      return null;
    }
    const directAdded = idsOf(direct).filter((id) => !before.has(id));
    if (directAdded.length === added.length) {
      added.forEach((id, index) => ids.set(id, directAdded[index]!));
    }
  }
  return direct;
};

const intentArbitrary = fc.record({
  // Structural kinds twice as often: they are the ones that compose badly.
  kind: fc.constantFrom(
    ...KINDS,
    "splitBlock",
    "mergeBlockWithNext",
    "insertAfterBlock",
    "deleteBlock",
    "insertTable",
  ),
  pick: fc.nat({ max: 50 }),
  detail: fc.nat({ max: 50 }),
  follow: fc.boolean(),
});

const caseArbitrary = fc.record({
  mode: fc.constantFrom(...MODES),
  intents: fc.array(intentArbitrary, { minLength: 1, maxLength: 5 }),
});

test("reject restores heading properties when an insertion interrupts a pending merge", async () => {
  const original = read(await open());
  const reviewer = await open();
  const heading = reviewer.getContent().at(0);
  const following = reviewer.getContent().at(1);
  if (!heading || !following)
    throw new Error("Missing heading and following paragraph");
  const edit = (operation: Operation) =>
    expect(apply(reviewer, "tracked-changes", operation)).toBe(true);
  edit({
    type: "insertAfterBlock",
    blockId: following.id,
    text: "Inserted clause.",
  });
  edit({ type: "deleteBlock", blockId: following.id });
  edit({
    type: "insertAfterBlock",
    blockId: following.id,
    text: "Inserted clause.",
  });
  edit({ type: "mergeBlockWithNext", blockId: heading.id, separator: " " });
  edit({
    type: "insertBeforeBlock",
    blockId: following.id,
    text: "Inserted clause.",
  });
  const saved = await reopen(reviewer);
  reviewer.rejectAll();
  expect(read(reviewer)).toEqual(original);
  expect(read(await reopen(reviewer))).toEqual(original);
  saved.rejectAll();
  expect(read(saved)).toEqual(original);
});

/**
 * The run goes right after the merged paragraph, whose words accepting runs
 * into the run's first paragraph (`insertionAfterPendingBreak`).
 */
for (const session of SESSIONS)
  test(`a run inserted after a pending merge resolves both ways without moving properties (${session})`, async () => {
    const original = read(await open(undefined, session));
    const edited = async () => {
      const reviewer = await open(undefined, session);
      const anchor = reviewer.getContent().at(1);
      if (!anchor) throw new Error("Missing merge anchor");
      expect(
        apply(reviewer, "tracked-changes", {
          type: "mergeBlockWithNext",
          blockId: anchor.id,
          separator: " ",
        }),
      ).toBe(true);
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            {
              id: "heading",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "A second-level heading.",
              styleId: "Heading2",
            },
            {
              id: "bullet",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "A new bullet.",
              numbering: { kind: "newList", format: "bullet" },
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(2);
      return reviewer;
    };

    const rejecting = await edited();
    rejecting.rejectAll();
    expect(read(rejecting)).toEqual(original);
    expect(read(await reopen(rejecting, session))).toEqual(original);

    const accepting = await edited();
    accepting.acceptAll();
    const accepted = accepting.getContent();
    const heading =
      session === "canonical"
        ? "This agreement is made between the parties named below. A second-level heading."
        : "A second-level heading.";
    expect(accepted.find(({ text }) => text === heading)?.kind).toBe("heading");
    expect(
      accepted
        .find(({ text }) => text.startsWith("A new bullet."))
        ?.text.startsWith("A new bullet."),
    ).toBe(true);
    expect(read(await reopen(accepting, session))).toEqual(read(accepting));
  });

for (const session of SESSIONS)
  for (const tableDeletion of ["deleteTable", "deleteTableRow"] as const) {
    test(`terminal paragraph deletions resolve like direct after ${tableDeletion} (${session})`, async () => {
      const original = read(await open(undefined, session));
      const edited = async (mode: Mode) => {
        const reviewer = await open(undefined, session);
        const blocks = reviewer.getContent();
        const cell = blocks.find(
          ({ table: cellTable }) => cellTable !== undefined,
        );
        if (!cell) throw new Error("Missing table deletion target");
        expect(
          apply(reviewer, mode, { type: tableDeletion, blockId: cell.id }),
        ).toBe(true);
        const joinedIds = new Map<string, string>();
        for (const prefix of ["Prices exclude", "Signed", "The Buyer"]) {
          const block = blocks.find(({ text }) => text.startsWith(prefix));
          if (!block)
            throw new Error("Missing terminal paragraph deletion target");
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [
                {
                  id: prefix,
                  type: "deleteBlock",
                  blockId: joinedIds.get(block.id) ?? block.id,
                },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied).toHaveLength(1);
          expect(result.skipped).toEqual([]);
          for (const { paragraphIdentity } of result.applied) {
            if (paragraphIdentity?.type === "joined")
              joinedIds.set(
                paragraphIdentity.retiredBlockId,
                paragraphIdentity.blockId,
              );
          }
        }
        return reviewer;
      };
      const direct = await edited("direct");
      const tracked = await edited("tracked-changes");
      const rejecting = await reopen(tracked, session);
      rejecting.rejectAll();
      expect(read(rejecting)).toEqual(original);
      const savedPending = await reopen(tracked, session);
      // On the canonical session the table pending deletion stands before the
      // last paragraphs when they are deleted: the last one keeps its mark and
      // loses its words, so accepting leaves one empty final paragraph more
      // than the direct edits (the stated table-before-final-paragraph shape).
      const accepted = [
        ...read(direct),
        ...(session === "canonical" && tableDeletion === "deleteTable"
          ? ["paragraph: "]
          : []),
      ];
      tracked.acceptAll();
      expect(read(tracked)).toEqual(accepted);
      expect(read(await reopen(tracked, session))).toEqual(accepted);
      savedPending.acceptAll();
      expect(read(savedPending)).toEqual(accepted);
    });
  }

test("reject removes tables inserted around an inserted terminal paragraph", async () => {
  const original = read(await open());
  const edit = async (mode: Mode) => {
    const reviewer = await open();
    const anchor = reviewer
      .getContent()
      .find(({ text }) => text === "Signed in two copies.");
    if (!anchor) throw new Error("Missing terminal paragraph anchor");
    for (const operation of [
      {
        type: "insertTable",
        blockId: anchor.id,
        rows: [["Term", "Value"]],
      },
      {
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: "Inserted clause.",
      },
      {
        type: "insertTable",
        blockId: anchor.id,
        rows: [["Term", "Value"]],
      },
    ] satisfies Operation[]) {
      expect(apply(reviewer, mode, operation)).toBe(true);
    }
    return reviewer;
  };

  const direct = await edit("direct");
  const tracked = await edit("tracked-changes");
  const savedPending = await reopen(tracked);
  tracked.rejectAll();
  expect(read(tracked)).toEqual(original);
  expect(read(await reopen(tracked))).toEqual(original);
  savedPending.rejectAll();
  expect(read(savedPending)).toEqual(original);

  const accepting = await edit("tracked-changes");
  accepting.acceptAll();
  expect(read(accepting)).toEqual(read(direct));
  expect(read(await reopen(accepting))).toEqual(read(direct));
});

/**
 * Tracked work around the original paragraphs: insertions, splits and
 * deletions of them and of what they became, and comments on them. A small
 * pick on a followed step aims at its target or the first block it added, so
 * steps compose on the same paragraphs; a run often inserts a paragraph,
 * splits it and deletes one of its halves.
 */
const identityIntent = (
  kind: fc.Arbitrary<Kind>,
  follow: fc.Arbitrary<boolean> = fc.constantFrom(true, true, false),
): fc.Arbitrary<Intent> =>
  fc.record({ kind, pick: fc.nat({ max: 2 }), detail: fc.nat(50), follow });

/** An insertion, a split of what it added and a deletion of one of its halves. */
const insertSplitDelete = fc.tuple(
  identityIntent(fc.constantFrom("insertBeforeBlock", "insertAfterBlock")),
  identityIntent(fc.constant("splitBlock"), fc.constant(true)),
  identityIntent(fc.constant("deleteBlock"), fc.constant(true)),
);

const anyIdentityIntent = fc.tuple(
  identityIntent(
    fc.constantFrom(
      "insertBeforeBlock",
      "insertAfterBlock",
      "splitBlock",
      "deleteBlock",
      "commentOnBlock",
    ),
  ),
);

const identityIntentsArbitrary = fc
  .array(
    fc.oneof(
      { weight: 2, arbitrary: insertSplitDelete },
      { weight: 1, arbitrary: anyIdentityIntent },
    ),
    { minLength: 1, maxLength: 3 },
  )
  .map((groups) => groups.flat());

const resolveAll = (
  reviewer: FolioDocxReviewer,
  decision: "reject" | "accept",
): void => {
  if (decision === "reject") reviewer.rejectAll();
  else reviewer.acceptAll();
};

const identityCaseArbitrary = fc.record({
  intents: identityIntentsArbitrary,
  saves: fc.array(fc.boolean(), { minLength: 1, maxLength: 10 }),
});

test("an insertion before an original paragraph, split, saved and half deleted resolves to the same ids after another save", async () => {
  const target = (await open()).getContent().at(1)?.id;
  if (target === undefined) throw new Error("Missing original paragraph");
  const edit = (reviewer: FolioDocxReviewer, operation: Operation) =>
    expect(apply(reviewer, "tracked-changes", operation)).toBe(true);
  for (const decision of ["reject", "accept"] as const) {
    const reviewer = await open();
    edit(reviewer, {
      type: "insertBeforeBlock",
      blockId: target,
      text: "Second added clause.",
    });
    const inserted = reviewer.getContent().at(1)?.id;
    if (inserted === undefined) throw new Error("Missing inserted paragraph");
    edit(reviewer, { type: "splitBlock", blockId: inserted, offset: 6 });
    const reopened = await reopen(reviewer);
    const half = reopened.getContent().at(2)?.id;
    if (half === undefined) throw new Error("Missing split half");
    edit(reopened, { type: "deleteBlock", blockId: half });
    const saved = await reopen(reopened);
    resolveAll(reopened, decision);
    resolveAll(saved, decision);
    expect(identities(saved)).toEqual(identities(reopened));
    if (decision === "reject")
      expect(idsOf(saved)).toEqual(idsOf(await open()));
  }
});

describe("resolving random tracked work", () => {
  test("resolving gives the same paragraph ids whether or not the document was saved on the way", async () => {
    await assertProperty(
      fc.asyncProperty(identityCaseArbitrary, async ({ intents, saves }) => {
        for (const decision of ["reject", "accept"] as const) {
          const live = await runSaving(intents, saves, false);
          const saved = await runSaving(intents, saves, true);
          const expected = () => identities(live.reviewer, live.minted);
          expect(identities(saved.reviewer, saved.minted)).toEqual(expected());
          // The edits made after the save reach a second save too.
          const savedAgain = await reopen(saved.reviewer);
          for (const reviewer of [live.reviewer, saved.reviewer, savedAgain])
            resolveAll(reviewer, decision);
          expect(identities(saved.reviewer, saved.minted)).toEqual(expected());
          expect(identities(savedAgain, saved.minted)).toEqual(expected());
        }
      }),
      {
        numRuns: 40,
        id: "resolving gives the same paragraph ids whether or not the document was saved on the way",
      },
    );
  });

  test("reject restores, accept equals direct, and both survive a save", async () => {
    const original = read(await open());
    const resolutionProperty = (counts: { comparedWithDirect: number }) =>
      fc.asyncProperty(caseArbitrary, async ({ mode, intents }) => {
        const reviewer = await open();
        const steps = run(reviewer, mode, intents);
        const pending = read(reviewer);

        // A pending tracked document survives its own save; a suggested one
        // saves without its suggestions, so it is only resolved live.
        let reopened: FolioDocxReviewer | null = null;
        if (mode === "tracked-changes") {
          const before = anchors(reviewer);
          reopened = await reopen(reviewer);
          expect(read(reopened)).toEqual(pending);
          expect(anchors(reopened)).toEqual(before);
        }

        // Rejecting: on a replay of the same intents, and on the reopened save.
        const replay = await open();
        run(replay, mode, intents);
        expect(read(replay)).toEqual(pending);
        replay.rejectAll();
        expect(read(replay)).toEqual(original);
        const savedRejected = await reopen(replay);
        expect(read(savedRejected)).toEqual(original);
        expect(anchors(savedRejected)).toEqual(anchors(replay));
        if (mode === "tracked-changes") {
          const rejecting = await reopen(reviewer);
          rejecting.rejectAll();
          expect(read(rejecting)).toEqual(original);
        }

        if (mode === "suggested") {
          expect(reviewer.acceptAll().resolved).toBe(0);
          expect(read(reviewer)).toEqual(pending);
          expect(read(await reopen(reviewer))).toEqual(original);
          return;
        }

        // Accepting: equal to the direct run, and the same after a save.
        const direct = await runDirectly(steps);
        reviewer.acceptAll();
        const accepted = read(reviewer);
        const acceptedAnchors = anchors(reviewer);
        const savedAccepted = await reopen(reviewer);
        expect(read(savedAccepted)).toEqual(accepted);
        expect(anchors(savedAccepted)).toEqual(acceptedAnchors);
        if (direct) {
          counts.comparedWithDirect++;
          expect(accepted).toEqual(read(direct));
        }
        if (reopened) {
          reopened.acceptAll();
          expect(read(reopened)).toEqual(accepted);
        }
      });
    // The seeds this property once failed on replay first, from
    // test/property-seeds/.
    await assertProperty(resolutionProperty({ comparedWithDirect: 0 }), {
      numRuns: 100,
      id: "reject restores, accept equals direct, and both survive a save",
    });
    // The direct comparison must not pass vacuously, on a fixed sample. Each
    // case opens and saves the document several times, so a fifth of the runs.
    const RUNS = 20;
    const counts = { comparedWithDirect: 0 };
    await checkCoverageSample(resolutionProperty(counts), RUNS);
    expect(counts.comparedWithDirect).toBeGreaterThan(0);
  });

  test("a comment reads the same across a save, whatever tracked edits follow it", async () => {
    await fc.assert(
      fc.asyncProperty(
        intentArbitrary,
        fc.array(intentArbitrary, { minLength: 1, maxLength: 3 }),
        async (comment, edits) => {
          const reviewer = await open();
          run(reviewer, "tracked-changes", [
            {
              ...comment,
              kind:
                comment.pick % 2 === 0 ? "commentOnRange" : "commentOnBlock",
              follow: false,
            },
            // Every edit aims at the commented paragraph or what it became.
            ...edits.map((edit) => ({ ...edit, follow: true })),
          ]);
          const survivesASave = async (
            state: FolioDocxReviewer,
          ): Promise<void> => {
            expect(anchors(await reopen(state))).toEqual(anchors(state));
          };
          await survivesASave(reviewer);
          const rejecting = await reopen(reviewer);
          rejecting.rejectAll();
          await survivesASave(rejecting);
          reviewer.acceptAll();
          await survivesASave(reviewer);
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  });
});
