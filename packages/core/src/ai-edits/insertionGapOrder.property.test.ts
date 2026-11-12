/**
 * Blocks a reviewer puts beside one anchor land in the order they were
 * requested, applied directly, as tracked changes accepted, and as
 * suggestions accepted and then accepted as tracked changes, whatever else
 * the batch does there.
 *
 * Every batch puts one to five blocks of mixed kinds (a paragraph, two
 * paragraphs from one text, a paragraph starting a new list, a table or a
 * signature table, whichever the mode has) after the anchor or before the
 * paragraph after it, which share one gap. The last batch may delete the
 * anchor too. Between batches the reviewer may save and reopen the package,
 * accept every change, or both.
 *
 * Within a batch the blocks keep the batch's order. A later batch after the
 * anchor lands right after it, ahead of what earlier batches put there; one
 * before the next paragraph lands right before it, behind them.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes" | "suggested";
type Kind = "paragraph" | "paragraphs" | "newList" | "table" | "signature";
type Between = "nothing" | "reopen" | "accept" | "reopenAndAccept";
type Batch = {
  side: "afterAnchor" | "beforeNext";
  kinds: Kind[];
  deletesAnchor: boolean;
};
type Case = {
  anchorIsLast: boolean;
  mode: Exclude<Mode, "direct">;
  batches: Batch[];
  between: Between[];
};

/** A block once accepted: its words, and whether a table holds it. */
type Block = { text: string; inTable: boolean };
/** What the requested order names: an insertion, or a paragraph that was there. */
type Placed = { label: string; table: boolean };

const OPENING = "Opening clause.";
const ANCHOR = "Anchor clause.";
const NEXT = "Next clause.";
const LABEL = /\bB\dI\d\b/u;

/**
 * Who put `block` there: the insertion its words name, or the paragraph it
 * is. A signature table's other cells (rules, captions) belong to the party
 * named before them.
 */
const labelsOf = (blocks: readonly Block[]): string[] => {
  const labels: string[] = [];
  for (const block of blocks) {
    const label =
      LABEL.exec(block.text)?.[0] ??
      (block.inTable ? labels.at(-1) : block.text);
    if (label !== undefined && label !== labels.at(-1)) labels.push(label);
  }
  return labels;
};

const insertionOf = (
  kind: Kind,
  label: string,
  anchor: { blockId: string; side: Batch["side"] },
): FolioDocumentOperation => {
  const id = `insert-${label}`;
  const position = anchor.side === "afterAnchor" ? "after" : "before";
  const beside = {
    type:
      anchor.side === "afterAnchor" ? "insertAfterBlock" : "insertBeforeBlock",
    blockId: anchor.blockId,
  } as const;
  switch (kind) {
    case "paragraph":
      return { id, ...beside, text: `${label} alone.`, alignment: "right" };
    case "paragraphs":
      return { id, ...beside, text: `${label} first.\n${label} second.` };
    case "newList":
      return {
        id,
        ...beside,
        text: `${label} listed.`,
        numbering: { kind: "newList", format: "numbered" },
      };
    case "table":
      return {
        id,
        type: "insertTable",
        blockId: anchor.blockId,
        position,
        rows: [[`${label} term`, `${label} value`]],
      };
    case "signature":
      return {
        id,
        type: "insertSignatureTable",
        blockId: anchor.blockId,
        position,
        parties: [{ name: `${label} signer` }],
      };
  }
};

const open = async (anchorIsLast: boolean): Promise<FolioDocxReviewer> => {
  const texts = [OPENING, ANCHOR, ...(anchorIsLast ? [] : [NEXT])];
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(texts.join("\n\n"))))
  ).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

const reopen = async (
  reviewer: FolioDocxReviewer,
): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });

/** The blocks a reader lists once every change is accepted and saved. */
const acceptedBlocks = async (
  reviewer: FolioDocxReviewer,
): Promise<Block[]> => {
  reviewer.acceptAll();
  return (await reopen(reviewer))
    .snapshot()
    .blocks.filter((block) => block.text.length > 0)
    .map((block) => ({ text: block.text, inTable: block.table !== undefined }));
};

/**
 * Replay `setup` in `mode`: the blocks accepting leaves, and the blocks the
 * requested order says, between the opening paragraph and the end.
 */
const replay = async (setup: Case, mode: Mode) => {
  let reviewer = await open(setup.anchorIsLast);
  const idOf = (text: string): string => {
    const block = reviewer.snapshot().blocks.find((b) => b.text === text);
    if (!block) throw new Error(`No block reads "${text}"`);
    return block.id;
  };
  const anchorId = idOf(ANCHOR);
  const nextId = setup.anchorIsLast ? undefined : idOf(NEXT);
  /** The insertions between the anchor and the next paragraph. */
  let expected: Placed[] = [];
  for (const [index, batch] of setup.batches.entries()) {
    const blockId = batch.side === "afterAnchor" ? anchorId : nextId;
    if (blockId === undefined) throw new Error("No paragraph after the anchor");
    const placed = batch.kinds.map((kind, at) => ({
      label: `B${index}I${at}`,
      table: kind === "table" || kind === "signature",
    }));
    const operations = batch.kinds.map((kind, at) =>
      insertionOf(kind, `B${index}I${at}`, { blockId, side: batch.side }),
    );
    if (batch.deletesAnchor) {
      operations.push({ id: "delete", type: "deleteBlock", blockId: anchorId });
    }
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        atomic: false,
        operations,
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    // A suggestion stays out of the package until accepted, which turns it
    // into an ordinary tracked change.
    const suggestionIds = new Set(
      result.applied.flatMap(({ suggestionId }) => suggestionId ?? []),
    );
    for (const suggestionId of suggestionIds) {
      expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
    }
    expected =
      batch.side === "afterAnchor"
        ? [...placed, ...expected]
        : [...expected, ...placed];
    const between = setup.between[index] ?? "nothing";
    if (between === "reopen" || between === "reopenAndAccept") {
      reviewer = await reopen(reviewer);
    }
    if (between === "accept" || between === "reopenAndAccept") {
      reviewer.acceptAll();
    }
  }
  const deleted = setup.batches.some((batch) => batch.deletesAnchor);
  return {
    accepted: await acceptedBlocks(reviewer),
    requested: [
      { label: OPENING, table: false },
      ...(deleted ? [] : [{ label: ANCHOR, table: false }]),
      ...expected,
      ...(setup.anchorIsLast ? [] : [{ label: NEXT, table: false }]),
    ],
  };
};

const kindsIn = (mode: Case["mode"]) =>
  fc.constantFrom<Kind[]>(
    "paragraph",
    "paragraphs",
    "newList",
    // A table has a tracked form and a signature table a suggested one.
    mode === "suggested" ? "signature" : "table",
  );

const caseArbitrary: fc.Arbitrary<Case> = fc
  .constantFrom<Case["mode"][]>("tracked-changes", "suggested")
  .chain((mode) =>
    fc.record({
      // A suggestion beside the story's last paragraph fails to accept, or
      // goes with a suggested deletion of that paragraph, which is not about
      // order: suggestions keep a paragraph after the anchor.
      anchorIsLast: mode === "suggested" ? fc.constant(false) : fc.boolean(),
      mode: fc.constant(mode),
    }),
  )
  .chain(({ anchorIsLast, mode }) =>
    fc
      .array(
        fc.record({
          side: anchorIsLast
            ? fc.constant<Batch["side"]>("afterAnchor")
            : fc.constantFrom<Batch["side"][]>("afterAnchor", "beforeNext"),
          kinds: fc.array(kindsIn(mode), { minLength: 1, maxLength: 5 }),
        }),
        { minLength: 1, maxLength: 3 },
      )
      .chain((batches) =>
        fc.record({
          anchorIsLast: fc.constant(anchorIsLast),
          mode: fc.constant(mode),
          batches: fc.boolean().map((deletes) =>
            batches.map((batch, index) => ({
              ...batch,
              deletesAnchor: deletes && index === batches.length - 1,
            })),
          ),
          between: fc.array(
            fc.constantFrom<Between[]>(
              "nothing",
              "reopen",
              "accept",
              "reopenAndAccept",
            ),
            { minLength: batches.length, maxLength: batches.length },
          ),
        }),
      ),
  );

/**
 * Two tables with nothing between them read as one table once saved, and a
 * story ends with a paragraph: neither is about order.
 */
const readsApart = (placed: readonly Placed[]): boolean =>
  placed.at(-1)?.table !== true &&
  placed.every((entry, index) => !entry.table || !placed[index + 1]?.table);

test(
  "blocks put beside one anchor land in requested order in every mode",
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const direct = await replay(setup, "direct");
        fc.pre(readsApart(direct.requested));
        expect(labelsOf(direct.accepted)).toEqual(
          direct.requested.map(({ label }) => label),
        );
        const reviewed = await replay(setup, setup.mode);
        expect(reviewed.accepted).toEqual(direct.accepted);
      }),
      {
        id: "blocks put beside one anchor land in requested order in every mode",
        numRuns: 80,
        examples: [
          // A paragraph after the story's last paragraph, pending; then a
          // paragraph and a table after it in a batch that deletes it.
          [
            {
              anchorIsLast: true,
              mode: "tracked-changes",
              batches: [
                {
                  side: "afterAnchor",
                  kinds: ["paragraph"],
                  deletesAnchor: false,
                },
                {
                  side: "afterAnchor",
                  kinds: ["paragraph", "table"],
                  deletesAnchor: true,
                },
              ],
              between: ["nothing", "nothing"],
            },
          ],
        ],
      },
    );
  },
  propertyTestTimeout(120_000),
);
