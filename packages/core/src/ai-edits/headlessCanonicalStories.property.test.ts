/**
 * A canonical headless reviewer edits headers, footers and notes on its
 * session, and resolves their changes with the body's.
 *
 * Each case draws a document a reviewer lists every story of, one of its
 * headers, footers or notes, and one public operation built from what a
 * reader of that story sees, then holds the edit to the laws the body's
 * edits keep:
 *
 * - undoing it gives back every read and the document exactly;
 * - tracked, then every change rejected, gives every story as rejecting
 *   every change without it gives;
 * - tracked or direct, a save and a reopen give back every story's text.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { packageDocumentArbitrary } from "../../../../test/generators/packageOperationArbitraries";
import {
  type OperationDraw,
  operationDrawArbitrary,
  operationFor,
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
import { type FolioDocumentStoryHandle, FolioDocxReviewer } from "./headless";

type Reviewer = FolioDocxReviewer;
type Mode = "direct" | "tracked-changes";
type StoryHandle = Exclude<FolioDocumentStoryHandle, { type: "main" }>;

const open = (buffer: ArrayBuffer): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Story Reviewer",
    session: "canonical",
  });

/** Every story's text, and the reads an undo must give back. */
const storyTexts = (reviewer: Reviewer) =>
  reviewer.listStories().map((story) => reviewer.readStory(story.handle));

const reads = (reviewer: Reviewer) => ({
  stories: storyTexts(reviewer),
  notes: reviewer.getNotesAsText(),
  content: reviewer.getContent(),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  document: reviewer.toDocument(),
});

const apply = (
  reviewer: Reviewer,
  story: StoryHandle,
  operation: FolioDocumentOperation,
  mode: Mode,
) =>
  reviewer.applyDocumentOperationsToStory({
    story,
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode,
    },
    undefinedReferences: "refuse",
  });

type StoryOutcome =
  | "edited"
  | "notAPackage"
  | "sessionRefused"
  | "noStory"
  | "noOperation"
  | "refusedOperation";

/** Every story once every change is rejected, or `null` when that refuses. */
const rejected = (reviewer: Reviewer) => {
  try {
    const { skipped } = reviewer.rejectAll();
    return skipped.length > 0 ? null : storyTexts(reviewer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return null;
    throw error;
  }
};

const storyLaws = async (
  buffer: ArrayBuffer,
  pick: number,
  draw: OperationDraw,
): Promise<StoryOutcome> => {
  let reviewer: Reviewer;
  try {
    reviewer = await open(buffer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const stories = reviewer
    .listStories()
    .flatMap(({ handle }) => (handle.type === "main" ? [] : [handle]));
  const story = stories.at(pick % Math.max(stories.length, 1));
  if (!story) return "noStory";
  const operation = operationFor(
    { getContent: () => reviewer.snapshotStory(story)?.blocks ?? [] },
    draw,
  );
  if (operation === null) return "noOperation";

  const before = reads(reviewer);
  const direct = apply(reviewer, story, operation, "direct");
  if (direct.applied.length === 0 || direct.undoHandle === null)
    return "refusedOperation";
  const edited = storyTexts(reviewer);
  expect(reviewer.undoDocumentOperations(direct.undoHandle)).toMatchObject({
    status: "undone",
  });
  expect(reads(reviewer)).toEqual(before);

  // The direct edit saves and reopens to the same stories.
  const saved = await open(buffer);
  apply(saved, story, operation, "direct");
  expect(storyTexts(saved)).toEqual(edited);
  expect(storyTexts(await open(await saved.toBuffer()))).toEqual(edited);

  const tracked = await open(buffer);
  if (apply(tracked, story, operation, "tracked-changes").applied.length === 0)
    return "refusedOperation";
  const trackedTexts = storyTexts(tracked);
  expect(storyTexts(await open(await tracked.toBuffer()))).toEqual(
    trackedTexts,
  );
  const rejectedEdit = rejected(tracked);
  const rejectedWithout = rejected(await open(buffer));
  if (rejectedEdit !== null && rejectedWithout !== null)
    expect(rejectedEdit).toEqual(rejectedWithout);
  return "edited";
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

const generatedStoryEdits = (tally: Map<StoryOutcome, number>) =>
  fc.asyncProperty(
    packageDocumentArbitrary,
    fc.nat(),
    operationDrawArbitrary,
    async (document, pick, draw) => {
      const buffer = await packaged(document);
      const outcome =
        buffer === null ? "notAPackage" : await storyLaws(buffer, pick, draw);
      tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
    },
  );

const GENERATED_RUNS = 30;
const COVERAGE_RUNS = 30;

describe("a canonical reviewer edits headers, footers and notes", () => {
  test(
    "generated story edits undo, save and reject as the laws say",
    async () => {
      await assertProperty(generatedStoryEdits(new Map()), {
        id: "generated story edits undo, save and reject as the laws say",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "most generated cases edit a story",
    async () => {
      const tally = new Map<StoryOutcome, number>();
      await checkCoverageSample(generatedStoryEdits(tally), COVERAGE_RUNS);
      expect(tally.get("edited") ?? 0).toBeGreaterThanOrEqual(
        COVERAGE_RUNS / 3,
      );
    },
    propertyTestTimeout(120_000),
  );
});
