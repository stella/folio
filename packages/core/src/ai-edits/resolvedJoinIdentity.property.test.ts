/**
 * A paragraph keeps its identity through accepting or rejecting the changes
 * a batch tracked: accepting leaves every paragraph the same batch applied
 * directly leaves, each under the id it had, and rejecting leaves the
 * document the batch started from, ids and all, whether every change is
 * resolved at once or one at a time.
 *
 * The batch deletes paragraphs (the story's last ones too, whose tracked
 * form removes the break before them and joins into what accepting empties)
 * and puts paragraphs before or after others (after the last one too, whose
 * tracked form moves the story's last mark onto the new paragraph). Each
 * original paragraph is the target of one operation at most.
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

type Kind = "delete" | "insertBefore" | "insertAfter";
type Edit = { kind: Kind; target: number };
type Resolution = "all" | "oneByOne";
type Case = { paragraphs: number; edits: Edit[]; resolution: Resolution };
type Mode = "direct" | "tracked-changes";
/** A paragraph as a reader lists it: its id and its words. */
type Listed = [id: string, text: string];

const textOf = (index: number) => `Clause ${String(index)}.`;

const open = async (paragraphs: number): Promise<FolioDocxReviewer> => {
  const markdown = Array.from({ length: paragraphs }, (_, index) =>
    textOf(index),
  ).join("\n\n");
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(markdown))))
    .docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

/** The paragraphs a reader lists once the package is saved and reopened. */
const listed = async (reviewer: FolioDocxReviewer): Promise<Listed[]> =>
  (
    await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "Reviewer",
    })
  )
    .snapshot()
    .blocks.map(({ id, text }): Listed => [id, text]);

/** The operation each edit is, on the paragraph `blockId` names. */
const OPERATIONS = {
  delete: (id: string, blockId: string): FolioDocumentOperation => ({
    id,
    type: "deleteBlock",
    blockId,
  }),
  insertBefore: (id: string, blockId: string): FolioDocumentOperation => ({
    id,
    type: "insertBeforeBlock",
    blockId,
    text: `Added ${id}.`,
  }),
  insertAfter: (id: string, blockId: string): FolioDocumentOperation => ({
    id,
    type: "insertAfterBlock",
    blockId,
    text: `Added ${id}.`,
  }),
} as const satisfies Record<
  Kind,
  (id: string, blockId: string) => FolioDocumentOperation
>;

const operationsOf = (
  edits: readonly Edit[],
  ids: readonly string[],
): FolioDocumentOperation[] =>
  edits.map(({ kind, target }, index) => {
    const blockId = ids[target];
    if (blockId === undefined) throw new Error(`No paragraph ${target}`);
    return OPERATIONS[kind](`edit-${String(index)}`, blockId);
  });

const applied = async (setup: Case, mode: Mode) => {
  const reviewer = await open(setup.paragraphs);
  const original = reviewer.snapshot().blocks.map(({ id }) => id);
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations: operationsOf(setup.edits, original),
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return { reviewer, original };
};

/** Resolve every tracked change `decision`'s way, at once or one at a time. */
const resolve = (
  reviewer: FolioDocxReviewer,
  decision: "accept" | "reject",
  resolution: Resolution,
) => {
  if (resolution === "all") {
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return;
  }
  for (
    let change = reviewer.getChanges().at(0);
    change !== undefined;
    change = reviewer.getChanges().at(0)
  ) {
    const resolved =
      decision === "accept"
        ? reviewer.acceptChange(change)
        : reviewer.rejectChange(change);
    expect(resolved).toBe(true);
  }
};

/** The listed paragraphs that were there before the batch. */
const originals = (blocks: readonly Listed[], ids: readonly string[]) =>
  blocks.filter(([id]) => ids.includes(id));

const caseArbitrary: fc.Arbitrary<Case> = fc
  .integer({ min: 2, max: 5 })
  .chain((paragraphs) =>
    fc.record({
      paragraphs: fc.constant(paragraphs),
      edits: fc.uniqueArray(
        fc.record({
          kind: fc.constantFrom<Kind[]>(
            "delete",
            "insertBefore",
            "insertAfter",
          ),
          target: fc.integer({ min: 0, max: paragraphs - 1 }),
        }),
        {
          minLength: 1,
          maxLength: paragraphs,
          selector: (edit) => edit.target,
        },
      ),
      resolution: fc.constantFrom<Resolution[]>("all", "oneByOne"),
    }),
  );

const examples: [Case][] = [
  // The story's last paragraph deleted: the break before it goes instead.
  [
    {
      paragraphs: 3,
      edits: [{ kind: "delete", target: 2 }],
      resolution: "all",
    },
  ],
  // The last two deleted, accepted one change at a time.
  [
    {
      paragraphs: 4,
      edits: [
        { kind: "delete", target: 3 },
        { kind: "delete", target: 2 },
      ],
      resolution: "oneByOne",
    },
  ],
  // Every paragraph deleted: the first is the one emptied and kept.
  [
    {
      paragraphs: 3,
      edits: [
        { kind: "delete", target: 0 },
        { kind: "delete", target: 1 },
        { kind: "delete", target: 2 },
      ],
      resolution: "all",
    },
  ],
  // A paragraph after the last one: rejecting it gives the last its own id back.
  [
    {
      paragraphs: 2,
      edits: [{ kind: "insertAfter", target: 1 }],
      resolution: "all",
    },
  ],
];

test(
  "accepting or rejecting tracked block edits keeps each paragraph's identity",
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const direct = await applied(setup, "direct");
        const directBlocks = await listed(direct.reviewer);

        // Every paragraph the direct edit keeps is there under its id, with
        // its words; a removed paragraph's mark may stay to end the story.
        const accepted = await applied(setup, "tracked-changes");
        resolve(accepted.reviewer, "accept", setup.resolution);
        const acceptedBlocks = await listed(accepted.reviewer);
        expect(
          originals(directBlocks, direct.original).filter(([id, text]) =>
            acceptedBlocks.some(
              ([acceptedId, acceptedText]) =>
                acceptedId === id && acceptedText === text,
            ),
          ),
        ).toEqual(originals(directBlocks, direct.original));

        const rejected = await applied(setup, "tracked-changes");
        resolve(rejected.reviewer, "reject", setup.resolution);
        expect(await listed(rejected.reviewer)).toEqual(
          rejected.original.map((id, index): Listed => [id, textOf(index)]),
        );
      }),
      {
        id: "accepting or rejecting tracked block edits keeps each paragraph's identity",
        numRuns: 60,
        examples,
      },
    );
  },
  propertyTestTimeout(60_000),
);
