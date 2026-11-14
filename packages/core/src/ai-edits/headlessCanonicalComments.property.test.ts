/**
 * A canonical headless reviewer replies to and resolves comments through the
 * session's semantic comment operations, with the legacy reviewer's
 * outcome.
 *
 * For every comment thread of a document, a reply is added and the thread
 * is resolved, then reopened, on both engines. The threads each engine reads
 * afterwards (author, text, anchored text, resolution and the replies'
 * authors and texts; ids are each engine's own) are the same, live and
 * after a save and a reopen.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { reviewerDocumentArbitrary } from "../../../../test/generators/reviewerDocumentArbitraries";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalCommentError } from "../canonical/canonicalComments";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import {
  CANONICAL_GAP,
  type HeadlessReviewerSession,
} from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const FIXTURES_DIR = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus",
);

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
};

const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const DATE = "2026-01-02T03:04:05Z";

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Comment Reviewer",
    session,
  });

/** The threads as a reader sees them, without the ids each engine mints. */
const threads = (reviewer: FolioDocxReviewer) =>
  reviewer.getComments().map((thread) => ({
    author: thread.author,
    text: thread.text,
    anchoredText: thread.anchoredText,
    blockId: thread.blockId,
    done: thread.done,
    replies: thread.replies.map(({ author, text, date }) => ({
      author,
      text,
      date,
    })),
  }));

/** Reply to every thread, resolve it, and reopen the first one. */
const review = (reviewer: FolioDocxReviewer) => {
  const roots = reviewer.getComments();
  for (const [index, thread] of roots.entries()) {
    const reply = reviewer.replyTo(thread, {
      text: `Reply ${String(index)}`,
      author: "Second Reviewer",
      date: DATE,
    });
    expect(reply).not.toBeNull();
    expect(reviewer.resolveComment(String(thread.id))).toBe(true);
  }
  const first = roots.at(0);
  if (first)
    expect(reviewer.resolveComment(String(first.id), { resolved: false })).toBe(
      true,
    );
  return roots.length;
};

type CommentOutcome =
  "reviewed" | "noThreads" | "sessionRefused" | "commentRefused";

/**
 * The session's typed comment refusal: a thread without a source anchor is
 * preserved read-only on the canonical engine.
 */
const isCommentRefusal = (error: unknown): boolean =>
  CanonicalCommentError.is(error) && error.gap === CANONICAL_GAP.comments;

const commentLaws = async (buffer: ArrayBuffer): Promise<CommentOutcome> => {
  let canonical: FolioDocxReviewer;
  try {
    canonical = await open(buffer, "canonical");
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const legacy = await open(buffer, "legacy");
  let reviewed: number;
  try {
    reviewed = review(canonical);
  } catch (error) {
    if (isCommentRefusal(error)) return "commentRefused";
    throw error;
  }
  if (reviewed === 0) return "noThreads";
  review(legacy);
  expect(threads(canonical)).toEqual(threads(legacy));
  const reopened = await open(await canonical.toBuffer(), "canonical");
  expect(threads(reopened)).toEqual(threads(canonical));
  // An unknown comment is neither replied to nor resolved.
  expect(canonical.replyTo(2_147_483_000, { text: "None" })).toBeNull();
  expect(canonical.resolveComment("2147483000")).toBe(false);
  return "reviewed";
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

const generatedComments = (tally: Map<string, number>) =>
  fc.asyncProperty(reviewerDocumentArbitrary, async (document) => {
    const buffer = await packaged(document);
    const outcome = buffer === null ? "notAPackage" : await commentLaws(buffer);
    tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
  });

const COMMENTED_FIXTURES: string[] = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .filter((name) => /comment/iu.test(name))
  .sort();

const GENERATED_RUNS = 25;
const COVERAGE_RUNS = 25;

describe("a canonical reviewer replies to and resolves comments", () => {
  test("commented corpus fixtures are present", () => {
    expect(COMMENTED_FIXTURES.length).toBeGreaterThan(0);
  });

  test.each(COMMENTED_FIXTURES)(
    "%s: replies and resolution read as on the legacy reviewer",
    async (filename) => {
      expect(await commentLaws(await ingested(readFixture(filename)))).not.toBe(
        "sessionRefused",
      );
    },
    propertyTestTimeout(30_000),
  );

  test(
    "generated threads read as on the legacy reviewer",
    async () => {
      await assertProperty(generatedComments(new Map()), {
        id: "generated threads read as on the legacy reviewer",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "some generated documents hold threads to review",
    async () => {
      const tally = new Map<string, number>();
      await checkCoverageSample(generatedComments(tally), COVERAGE_RUNS);
      expect(tally.get("reviewed") ?? 0).toBeGreaterThan(0);
    },
    propertyTestTimeout(120_000),
  );
});

describe("a comment on words a tracked replacement removes", () => {
  for (const mode of ["tracked-changes", "suggested"] as const) {
    test(`stays on the removed words (${mode}), as on the legacy reviewer`, async () => {
      const buffer = await ingested(
        await createDocx(
          fromMarkdown(
            "The Supplier delivers the goods on time.\n\nThe Buyer pays.",
          ),
        ),
      );
      const engines = await Promise.all(
        (["legacy", "canonical"] as const).map(async (session) => {
          const reviewer = await open(buffer, session);
          const block = reviewer
            .snapshotStory({ type: "main" })
            ?.blocks.find(({ kind }) => kind === "paragraph");
          if (block?.kind !== "paragraph") throw new Error("no paragraph");
          const run = (
            operations: Parameters<
              FolioDocxReviewer["applyDocumentOperations"]
            >[0]["operations"],
            batchMode: "direct" | "tracked-changes" | "suggested",
          ) =>
            reviewer.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                operations,
                mode: batchMode,
              },
              { undefinedReferences: "refuse" },
            ).applied.length;
          for (const quote of ["time", "The"])
            expect(
              run(
                [
                  {
                    type: "commentOnBlock",
                    blockId: block.id,
                    quote,
                    comment: { text: `On ${quote}.` },
                  },
                ],
                "tracked-changes",
              ),
            ).toBe(1);
          expect(
            run(
              [
                {
                  type: "replaceBlock",
                  blockId: block.id,
                  text: "Invoice delivery payment agreement supplier buyer payment term.",
                },
              ],
              mode,
            ),
          ).toBe(1);
          const texts = (read: FolioDocxReviewer) =>
            threads(read)
              .map(({ text }) => text)
              .sort();
          return {
            live: texts(reviewer),
            saved: texts(await open(await reviewer.toBuffer(), session)),
          };
        }),
      );
      const [legacy, canonical] = engines;
      expect(legacy?.live).toEqual(["On The.", "On time."]);
      expect(canonical).toEqual(legacy);
    });
  }
});
