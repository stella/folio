/**
 * A canonical headless reviewer stages, exports, restages and resolves
 * pending suggestions on its session.
 *
 * Each case draws a document and one public operation built from what a
 * reader sees, applied as a suggestion:
 *
 * - the saved package holds the suggestion as rejected: reopened, it reads
 *   as the document before the suggestion;
 * - the exported records restage on a reviewer reopened on the saved
 *   package, which then reads as the suggesting reviewer did;
 * - rejecting the suggestion gives the document before it back;
 * - accepting it adopts it as ordinary tracked changes: accepting every
 *   change then gives the words of the same operation applied directly.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  type OperationDraw,
  operationFor,
  stagingDrawArbitrary,
} from "../../../../test/generators/reviewerOperationDraws";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

type Reviewer = FolioDocxReviewer;

const open = (buffer: ArrayBuffer): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Suggesting Reviewer",
    session: "canonical",
  });

const apply = (
  reviewer: Reviewer,
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes" | "suggested",
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode,
    },
    { undefinedReferences: "refuse" },
  );

/** What a reader sees, pending suggestions included. */
const reads = (reviewer: Reviewer) => ({
  text: reviewer.getContentAsText({ annotated: true }),
  changes: reviewer.getChanges().map(({ type, text, author, blockId }) => ({
    type,
    text,
    author,
    blockId,
  })),
  comments: reviewer
    .getComments()
    .map(({ text, anchoredText, author }) => ({ text, anchoredText, author })),
});

const words = (text: string): string => text.replace(/\s+/gu, "");

type SuggestionOutcome =
  | "suggested"
  | "notAPackage"
  | "sessionRefused"
  | "noOperation"
  | "refusedOperation"
  | "resolutionRefused";

const suggestionLaws = async (
  buffer: ArrayBuffer,
  draw: OperationDraw,
): Promise<SuggestionOutcome> => {
  let reviewer: Reviewer;
  try {
    reviewer = await open(buffer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const operation = operationFor(reviewer, draw, {
    replaceIn: "blocksWithWords",
  });
  if (operation === null) return "noOperation";
  const before = reads(reviewer);
  const result = apply(reviewer, operation, "suggested");
  if (result.applied.length === 0) return "refusedOperation";
  const suggested = reads(reviewer);
  const records = reviewer.exportPendingSuggestions();
  expect(records.length).toBeGreaterThan(0);
  const suggestionIds = [
    ...new Set(records.map(({ suggestionId }) => suggestionId)),
  ];

  // The file holds the suggestion as rejected.
  const saved = await reviewer.toBuffer();
  expect(reads(await open(saved))).toEqual(before);

  // The exported records restage on the saved package.
  const restaged = await open(saved);
  for (const loaded of restaged.loadPendingSuggestions(records))
    expect(loaded).toMatchObject({ status: "restaged" });
  expect(reads(restaged)).toEqual(suggested);

  // Rejecting the suggestion gives the document back.
  for (const id of suggestionIds)
    expect(restaged.rejectSuggestion(id)).toBe(true);
  expect(reads(restaged)).toEqual(before);

  // Accepting adopts it; accepting every change then is the direct edit.
  for (const id of suggestionIds)
    expect(reviewer.acceptSuggestion(id)).toBe(true);
  expect(reviewer.exportPendingSuggestions()).toEqual([]);
  const direct = await open(buffer);
  apply(direct, operation, "direct");
  try {
    const accepted = reviewer.acceptAll();
    const directAccepted = direct.acceptAll();
    if (accepted.skipped.length > 0 || directAccepted.skipped.length > 0)
      return "resolutionRefused";
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "resolutionRefused";
    throw error;
  }
  expect(words(reviewer.getContentAsText())).toBe(
    words(direct.getContentAsText()),
  );
  return "suggested";
};

const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

const generatedSuggestions = (tally: Map<SuggestionOutcome, number>) =>
  fc.asyncProperty(
    reviewerDocumentArbitrary,
    stagingDrawArbitrary,
    async (document, draw) => {
      const buffer = await packaged(document);
      const outcome =
        buffer === null ? "notAPackage" : await suggestionLaws(buffer, draw);
      tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
    },
  );

const GENERATED_RUNS = 30;
const COVERAGE_RUNS = 30;

describe("a canonical reviewer's pending suggestions", () => {
  test(
    "generated suggestions save, restage, reject and accept as the laws say",
    async () => {
      await assertProperty(generatedSuggestions(new Map()), {
        id: "generated suggestions save, restage, reject and accept as the laws say",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "most generated cases stage a suggestion",
    async () => {
      const tally = new Map<SuggestionOutcome, number>();
      await checkCoverageSample(generatedSuggestions(tally), COVERAGE_RUNS);
      // The whole tally is in the failure, so a shortfall names its outcomes.
      expect({
        staged: (tally.get("suggested") ?? 0) >= COVERAGE_RUNS / 3,
        tally: Object.fromEntries(tally),
      }).toMatchObject({ staged: true });
    },
    propertyTestTimeout(120_000),
  );
});
