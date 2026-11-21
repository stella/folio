/**
 * A batch of formatting equals its operations applied one at a time in input
 * order, on both sessions and in every mode: a later write of a property
 * replaces an earlier one wherever their ranges meet, whether they start at
 * the same offset or not. The batch and the one-at-a-time run save the same
 * package, up to the revision ids each allocates, and skip the same
 * operations for the same reasons.
 *
 * A case opens and saves two packages, a few tens of milliseconds; the budget
 * is 60 s for the stated runs, scaled by the nightly factor.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIInlineFormattingPatch } from "./types";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const MODES = [
  "direct",
  "tracked-changes",
  "suggested",
] as const satisfies readonly FolioDocumentOperationMode[];

const TEXT = "Formatting target words";

const PATCHES = [
  { bold: true },
  { bold: null },
  { italic: true },
  { italic: null },
] as const satisfies readonly FolioAIInlineFormattingPatch[];

type Patch = (typeof PATCHES)[number];

/** One `formatRange` over `[start, end)` of the paragraph. */
type Format = { start: number; end: number; formatting: Patch };

type Case = {
  mode: FolioDocumentOperationMode;
  formats: readonly Format[];
};

const source = async (): Promise<ArrayBuffer> => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      content: [{ type: "run", content: [{ type: "text", text: TEXT }] }],
    },
  ];
  return await createDocx(document);
};

const rangeArbitrary = (start: fc.Arbitrary<number>) =>
  start.chain((from) =>
    fc.record({
      start: fc.constant(from),
      end: fc.integer({ min: from + 1, max: TEXT.length }),
      formatting: fc.constantFrom(...PATCHES),
    }),
  );

/** Two to four formats, all starting at one offset, or anywhere. */
const caseArbitrary: fc.Arbitrary<Case> = fc.record({
  mode: fc.constantFrom(...MODES),
  formats: fc.oneof(
    fc.integer({ min: 0, max: TEXT.length - 1 }).chain((start) =>
      fc.array(rangeArbitrary(fc.constant(start)), {
        minLength: 2,
        maxLength: 4,
      }),
    ),
    fc.array(rangeArbitrary(fc.integer({ min: 0, max: TEXT.length - 1 })), {
      minLength: 2,
      maxLength: 4,
    }),
  ),
});

const DATE = "2026-01-02T03:04:05.000Z";
/** Each batch allocates above the ids the batches before it took. */
const ID_SEED_STEP = 100;

type RunOptions = {
  session: HeadlessReviewerSession;
  mode: FolioDocumentOperationMode;
  batches: readonly (readonly Format[])[];
};

/** The saved body, with revision ids renumbered in order of appearance. */
const savedBody = async (reviewer: FolioDocxReviewer): Promise<string> => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? panic("Expected word/document.xml");
  const renumbered = new Map<string, number>();
  return xml.replaceAll(/\bw:id="(\d+)"/gu, (_, id: string) => {
    const next = renumbered.get(id) ?? renumbered.size;
    renumbered.set(id, next);
    return `w:id="${next}"`;
  });
};

/** Apply each batch in turn; what the package saves and what was skipped. */
const run = async ({ session, mode, batches }: RunOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
    author: "Reviewer",
    session,
  });
  const block = reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
  const all = batches
    .flat()
    .map(({ start, end, formatting }, index): FolioDocumentOperation => ({
      id: `format-${index}`,
      type: "formatRange",
      range:
        createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: start,
          endOffset: end,
        }) ?? panic("Expected a range"),
      formatting,
    }));
  const skipped: { id: string; reason: string }[] = [];
  let taken = 0;
  for (const [batchIndex, formats] of batches.entries()) {
    const operations = all.slice(taken, taken + formats.length);
    taken += formats.length;
    const result = reviewer.applyDocumentOperations(
      { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
      {
        undefinedReferences: "refuse",
        revisionStamp: {
          date: DATE,
          idSeed: ID_SEED_STEP * (batchIndex + 1),
        },
      },
    );
    for (const { id, reason } of result.skipped) skipped.push({ id, reason });
  }
  return { body: await savedBody(reviewer), skipped };
};

/** Formats that start at one offset and write one property both ways. */
const examples: [Case][] = MODES.flatMap((mode) => [
  [
    {
      mode,
      formats: [
        { start: 0, end: 10, formatting: { bold: true } },
        { start: 0, end: 10, formatting: { bold: null } },
      ],
    },
  ],
  [
    {
      mode,
      formats: [
        { start: 0, end: 10, formatting: { bold: true } },
        { start: 2, end: 6, formatting: { bold: null } },
      ],
    },
  ],
]);

describe.each(SESSIONS)("%s: a batch of formatting", (session) => {
  test(
    "equals its operations applied one at a time in input order",
    async () => {
      await assertProperty(
        fc.asyncProperty(caseArbitrary, async ({ mode, formats }) => {
          const batch = await run({ session, mode, batches: [formats] });
          const oneAtATime = await run({
            session,
            mode,
            batches: formats.map((format) => [format]),
          });
          expect(batch).toEqual(oneAtATime);
        }),
        {
          id: `${session}: a batch of formatting equals its operations one at a time`,
          numRuns: 40,
          examples,
        },
      );
    },
    propertyTestTimeout(60_000),
  );
});

describe("bold, then bold cleared, over one word in one batch", () => {
  test("canonical: the word ends as it started", async () => {
    const { body, skipped } = await run({
      session: "canonical",
      mode: "tracked-changes",
      batches: [
        [
          { start: 0, end: 10, formatting: { bold: true } },
          { start: 0, end: 10, formatting: { bold: null } },
        ],
      ],
    });
    expect(skipped).toEqual([]);
    expect(body).not.toContain("<w:b/>");
    expect(body).not.toContain("w:rPrChange");
  });

  test("legacy: the clearing is refused over the pending change the batch made", async () => {
    const { body, skipped } = await run({
      session: "legacy",
      mode: "tracked-changes",
      batches: [
        [
          { start: 0, end: 10, formatting: { bold: true } },
          { start: 0, end: 10, formatting: { bold: null } },
        ],
      ],
    });
    expect(skipped).toEqual([
      { id: "format-1", reason: "pendingRunPropertyChange" },
    ]);
    expect(body).toContain("<w:b/>");
  });
});
