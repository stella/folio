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
 * - `acceptAll` gives what the same operations give in `direct` mode, when
 *   each of them has a direct counterpart (its target exists there too);
 * - the resolved document survives a save: reopened, it reads the same;
 * - a pending (tracked) document saved and reopened resolves to the same
 *   result both ways as the live one;
 * - a comment's anchored text reads the same before and after a save.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph, table } from "../docx/server/build";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  isFolioDocumentOperationModeSupported,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { createFolioAITextRangeHandle } from "./snapshot";
import { FolioDocxReviewer } from "./headless";

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

const SENTENCES = ["Inserted clause.", "Payment is due.", "New term here.", "Short."] as const;

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
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  base = docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer;
  return base;
};

const open = async (bytes?: ArrayBuffer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(bytes ?? (await baseDocument()), { author: "Reviewer" });

type Block = ReturnType<FolioDocxReviewer["getContent"]>[number];

/** What a reader is left with: every block's kind and text. */
const read = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map((block) => `${block.kind}: ${block.text}`);

const anchors = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getComments().map((comment) => comment.anchoredText);

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  open(await reviewer.toBuffer());

const words = (text: string) =>
  [...text.matchAll(/[A-Za-z]{3,}/gu)].map((match) => ({ word: match[0], start: match.index }));

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
    (follow ? from.find(({ id }) => id === focus[pick % Math.max(focus.length, 1)]) : undefined) ??
    from[pick % Math.max(from.length, 1)];
  const sentence = SENTENCES[detail % SENTENCES.length];
  switch (kind) {
    case "insertAfterBlock":
    case "insertBeforeBlock": {
      const block = choose(body);
      return block && { type: kind, blockId: block.id, text: sentence };
    }
    case "deleteBlock": {
      const block = choose(body);
      return body.length > 2 && block ? { type: kind, blockId: block.id } : undefined;
    }
    case "replaceBlock": {
      const block = choose(body);
      return block && { type: kind, blockId: block.id, text: sentence };
    }
    case "replaceInBlock": {
      const block = choose(worded);
      const found = block && words(block.text)[detail % words(block.text).length];
      return (
        block && found && { type: kind, blockId: block.id, find: found.word, replace: sentence }
      );
    }
    case "splitBlock": {
      const block = choose(worded.filter((candidate) => words(candidate.text).length > 1));
      const at = block && words(block.text).slice(1)[detail % (words(block.text).length - 1)];
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
      return block && { type: kind, blockId: block.id, rows: [["Term", "Value"]] };
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
      const quote = block && words(block.text)[detail % words(block.text).length];
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
      const found = block && words(block.text)[detail % words(block.text).length];
      const range =
        block &&
        found &&
        createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: found.start,
          endOffset: found.start + found.word.length,
        });
      return range ? { type: kind, range, comment: { text: "Why?" } } : undefined;
    }
  }
};

type Operation = NonNullable<ReturnType<typeof operationFor>>;

const apply = (reviewer: FolioDocxReviewer, mode: Mode, operation: Operation): boolean =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [{ id: "1", ...operation }],
  } as never).applied.length === 1;

const idsOf = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map((block) => block.id);

/**
 * An applied operation, the text its target read, whether a table followed
 * the target, and the blocks it added.
 */
type Step = { operation: Operation; targetText: string; beforeTable: boolean; added: string[] };

const targetOf = (operation: Operation): string =>
  "blockId" in operation ? operation.blockId : operation.range.blockId;

/** Apply each intent the mode supports; what each applied step added. */
const run = (reviewer: FolioDocxReviewer, mode: Mode, intents: readonly Intent[]): Step[] => {
  const steps: Step[] = [];
  let focus: string[] = reviewer
    .getContent()
    .slice(-1)
    .map(({ id }) => id);
  for (const intent of intents) {
    if (!isFolioDocumentOperationModeSupported(intent.kind, mode)) {
      continue;
    }
    const before = new Set(idsOf(reviewer));
    const blocks = reviewer.getContent();
    const operation = operationFor(blocks, intent, focus);
    const index = operation ? blocks.findIndex(({ id }) => id === targetOf(operation)) : -1;
    const targetText = blocks[index]?.text;
    const beforeTable = blocks[index + 1]?.table !== undefined;
    if (operation && targetText !== undefined && apply(reviewer, mode, operation)) {
      const added = idsOf(reviewer).filter((id) => !before.has(id));
      steps.push({ operation, targetText, beforeTable, added });
      const after = reviewer.getContent();
      const at = after.findIndex(({ id }) => id === targetOf(operation));
      focus = [
        targetOf(operation),
        ...added,
        ...[after[at - 1], after[at + 1], after.at(-1)].flatMap((block) =>
          block ? [block.id] : [],
        ),
      ];
    }
  }
  return steps;
};

/**
 * The same steps in `direct` mode, or `null` when one has no counterpart
 * there: its target is gone (a block only a pending deletion still shows), or
 * direct mode refuses it. New block ids are random, so the blocks each step
 * adds are paired in order.
 */
const runDirectly = async (steps: readonly Step[]): Promise<FolioDocxReviewer | null> => {
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

let comparedWithDirect = 0;

describe("resolving random tracked work", () => {
  test("reject restores, accept equals direct, and both survive a save", async () => {
    const original = read(await open());
    const resolutionProperty = fc.asyncProperty(caseArbitrary, async ({ mode, intents }) => {
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

      // Accepting: equal to the direct run, and the same after a save.
      const direct = await runDirectly(steps);
      reviewer.acceptAll();
      const accepted = read(reviewer);
      const acceptedAnchors = anchors(reviewer);
      const savedAccepted = await reopen(reviewer);
      expect(read(savedAccepted)).toEqual(accepted);
      expect(anchors(savedAccepted)).toEqual(acceptedAnchors);
      if (direct) {
        comparedWithDirect++;
        expect(accepted).toEqual(read(direct));
      }
      if (reopened) {
        reopened.acceptAll();
        expect(read(reopened)).toEqual(accepted);
      }
    });
    await fc.assert(resolutionProperty, propertyConfig({ numRuns: 100, seed: 2055257210 }));
    await fc.assert(resolutionProperty, propertyConfig({ numRuns: 100, seed: -1401551044 }));
    await fc.assert(resolutionProperty, propertyConfig({ numRuns: 100, seed: -304239210 }));
    await fc.assert(resolutionProperty, propertyConfig({ numRuns: 100 }));
    // The direct comparison must not pass vacuously.
    expect(comparedWithDirect).toBeGreaterThan(0);
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
              kind: comment.pick % 2 === 0 ? "commentOnRange" : "commentOnBlock",
              follow: false,
            },
            // Every edit aims at the commented paragraph or what it became.
            ...edits.map((edit) => ({ ...edit, follow: true })),
          ]);
          const survivesASave = async (state: FolioDocxReviewer): Promise<void> => {
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
