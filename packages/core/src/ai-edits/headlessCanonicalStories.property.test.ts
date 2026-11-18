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

import { afterEach, describe, expect, setSystemTime, test } from "bun:test";
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
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
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

const storyChanges = (reviewer: Reviewer) =>
  reviewer
    .listStories()
    .flatMap(
      ({ handle }) =>
        reviewer.readReviewedStory({ story: handle, view: "current-markup" })
          ?.changes ?? [],
    );

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
  | "refusedOperation"
  | "resolutionRefused";

/** Every story once every change is rejected, or `null` when that refuses. */
const rejected = (reviewer: Reviewer) => {
  try {
    const changes = storyChanges(reviewer).length;
    const { resolved, skipped } = reviewer.rejectAll();
    if (skipped.length > 0) return null;
    expect(resolved).toBe(changes);
    expect(storyChanges(reviewer)).toEqual([]);
    return storyTexts(reviewer);
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
  expect(apply(saved, story, operation, "direct").applied).toHaveLength(1);
  expect(storyTexts(saved)).toEqual(edited);
  expect(storyTexts(await open(await saved.toBuffer()))).toEqual(edited);

  const tracked = await open(buffer);
  if (apply(tracked, story, operation, "tracked-changes").applied.length === 0)
    return "refusedOperation";
  const trackedTexts = storyTexts(tracked);
  expect(storyTexts(await open(await tracked.toBuffer()))).toEqual(
    trackedTexts,
  );

  // The accepted edit and the direct edit resolve the whole package alike.
  const accepted = tracked.fork();
  const directAccepted = saved.fork();
  try {
    const acceptedChanges = storyChanges(accepted).length;
    const directChanges = storyChanges(directAccepted).length;
    const acceptedResult = accepted.acceptAll();
    const directResult = directAccepted.acceptAll();
    if (acceptedResult.skipped.length > 0 || directResult.skipped.length > 0)
      return "resolutionRefused";
    expect(acceptedResult.resolved).toBe(acceptedChanges);
    expect(directResult.resolved).toBe(directChanges);
    expect(storyChanges(accepted)).toEqual([]);
    expect(storyChanges(directAccepted)).toEqual([]);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "resolutionRefused";
    throw error;
  }
  expect(storyTexts(accepted)).toEqual(storyTexts(directAccepted));

  const rejectedEdit = rejected(tracked);
  const rejectedWithout = rejected(await open(buffer));
  if (rejectedEdit === null || rejectedWithout === null)
    return "resolutionRefused";
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

afterEach(() => setSystemTime());

const REVIEW_STORIES = [
  { type: "main" },
  { type: "header", relationshipId: "rIdReviewHeader" },
  { type: "footer", relationshipId: "rIdReviewFooter" },
  { type: "footnote", noteId: 7 },
  { type: "endnote", noteId: 8 },
] as const;

/** One authored paragraph of every story kind, with live note references. */
const reviewPackage = async () => {
  const document = fromMarkdown("Body words.");
  document.package.document.content.push(
    paragraph([
      {
        type: "run",
        content: [
          { type: "footnoteRef", id: 7 },
          { type: "endnoteRef", id: 8 },
        ],
      },
    ]),
  );
  document.package.document.finalSectionProperties = {
    headerReferences: [{ type: "default", rId: "rIdReviewHeader" }],
    footerReferences: [{ type: "default", rId: "rIdReviewFooter" }],
  };
  document.package.headers = new Map([
    [
      "rIdReviewHeader",
      {
        type: "header",
        hdrFtrType: "default",
        content: [paragraph([run("Header words.")])],
      },
    ],
  ]);
  document.package.footers = new Map([
    [
      "rIdReviewFooter",
      {
        type: "footer",
        hdrFtrType: "default",
        content: [paragraph([run("Footer words.")])],
      },
    ],
  ]);
  document.package.footnotes = [
    { type: "footnote", id: 7, content: [paragraph([run("Footnote words.")])] },
  ];
  document.package.endnotes = [
    { type: "endnote", id: 8, content: [paragraph([run("Endnote words.")])] },
  ];
  return ingested(await createDocx(document));
};

for (const story of REVIEW_STORIES)
  for (const view of ["original", "final"] as const) {
    test(`${story.type} ${view} story review refuses its staged suggestion without changing state or history`, async () => {
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const reviewer = await open(await reviewPackage());
      const beforeStage = reads(reviewer);
      const block = reviewer.snapshotStory(story)?.blocks.at(0);
      if (!block) throw new Error("The review fixture has no target block");
      const staged = reviewer.applyDocumentOperationsToStory({
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "suggested",
          operations: [
            {
              id: "stage",
              type: "replaceBlock",
              blockId: block.id,
              text: "Proposed words.",
            },
          ],
        },
        undefinedReferences: "refuse",
      });
      expect(staged.applied).toHaveLength(1);
      expect(staged.skipped).toEqual([]);
      if (staged.undoHandle === null)
        throw new Error("The stage has no undo handle");
      const before = reads(reviewer);
      const records = reviewer.exportPendingSuggestions();
      expect(records).toHaveLength(1);
      const saved = new Uint8Array(await reviewer.toBuffer());
      expect(() => reviewer.resolveReviewedStory({ story, view })).toThrow(
        CanonicalSessionError,
      );
      expect(reads(reviewer)).toEqual(before);
      expect(reviewer.exportPendingSuggestions()).toEqual(records);
      expect(new Uint8Array(await reviewer.toBuffer())).toEqual(saved);

      // A different story is independently reviewable beside the staged one.
      const other = REVIEW_STORIES.find(
        (candidate) => candidate.type !== story.type,
      );
      if (!other) throw new Error("The review fixture has no other story");
      expect(reviewer.resolveReviewedStory({ story: other, view })).toBe(true);
      expect(reads(reviewer)).toEqual(before);
      expect(reviewer.exportPendingSuggestions()).toEqual(records);

      // Reading a resolved view may refuse, but cannot resolve the live stage.
      try {
        reviewer.readReviewedStory({ story, view });
      } catch (error) {
        if (!CanonicalSessionError.is(error)) throw error;
      }
      expect(reads(reviewer)).toEqual(before);
      expect(reviewer.exportPendingSuggestions()).toEqual(records);
      expect(reviewer.undoDocumentOperations(staged.undoHandle)).toMatchObject({
        status: "undone",
      });
      expect(reads(reviewer)).toEqual(beforeStage);
      expect(reviewer.exportPendingSuggestions()).toEqual([]);
    });

    test(`${story.type} ${view} story review resolves ordinary tracked edits`, async () => {
      const reviewer = await open(await reviewPackage());
      const block = reviewer.snapshotStory(story)?.blocks.at(0);
      if (!block) throw new Error("The review fixture has no target block");
      const before = reviewer.readStory(story)?.text;
      const result = reviewer.applyDocumentOperationsToStory({
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            {
              id: "edit",
              type: "replaceBlock",
              blockId: block.id,
              text: "Edited words.",
            },
          ],
        },
        undefinedReferences: "refuse",
      });
      expect(result.applied).toHaveLength(1);
      expect(result.skipped).toEqual([]);
      expect(reviewer.resolveReviewedStory({ story, view })).toBe(true);
      expect(
        reviewer.readReviewedStory({ story, view: "current-markup" })?.changes,
      ).toEqual([]);
      const text = reviewer.readStory(story)?.text;
      if (view === "original") expect(text).toBe(before);
      else expect(text).toContain("Edited words.");
      expect(reviewer.exportPendingSuggestions()).toEqual([]);
    });
  }

for (const session of ["legacy", "canonical"] as const) {
  test(`${session} forks keep story states and comment overrides independent`, async () => {
    const original = await FolioDocxReviewer.fromBuffer(await reviewPackage(), {
      session,
    });
    const header = REVIEW_STORIES[1];
    const block = original.snapshotStory(header)?.blocks.at(0);
    if (!block) throw new Error("The fixture has no header block");
    const editHeader = (reviewer: Reviewer, text: string) =>
      reviewer.applyDocumentOperationsToStory({
        story: header,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [
            { id: "header", type: "replaceBlock", blockId: block.id, text },
          ],
        },
        undefinedReferences: "refuse",
      });
    expect(editHeader(original, "Authored header.").applied).toHaveLength(1);
    const body = original.getContent().at(0);
    if (!body) throw new Error("The fixture has no body block");
    expect(
      original.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [
            {
              id: "comment",
              type: "commentOnBlock",
              blockId: body.id,
              quote: "Body",
              comment: { text: "Review body." },
            },
          ],
        },
        { undefinedReferences: "refuse" },
      ).applied,
    ).toHaveLength(1);
    const comment = original.getComments().at(0);
    if (!comment) throw new Error("The comment operation made no comment");
    expect(
      original.replyTo(comment, { text: "Authored reply." }),
    ).not.toBeNull();
    expect(original.resolveComment(String(comment.id))).toBe(true);
    const before = reads(original);
    const copy = original.fork();
    expect(reads(copy)).toEqual(before);
    expect(editHeader(copy, "Copy header.").applied).toHaveLength(1);
    expect(copy.resolveComment(String(comment.id), { resolved: false })).toBe(
      true,
    );
    expect(copy.replyTo(comment.id, { text: "Copy reply." })).not.toBeNull();
    expect(reads(original)).toEqual(before);
    const copied = reads(copy);
    expect(editHeader(original, "Original header.").applied).toHaveLength(1);
    expect(
      original.replyTo(comment.id, { text: "Original reply." }),
    ).not.toBeNull();
    expect(reads(copy)).toEqual(copied);
  });
}
