import { panic } from "better-result";
/** L8 gates fidelity; L9 measures differences between independent resolvers. */
import { afterAll, describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import type { Document } from "@stll/docx-core/model";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  normalizeForOps,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import {
  at,
  findParagraph,
  FIRST_ID,
  fixture,
  newIds,
  operations,
  seedArbitrary,
  stamp,
} from "../../../typecheck/ops/reviewGenerators.typecheck";
import {
  compareReviewResolution,
  editorProjectionRoundTrip,
  editorRoundTrip,
  reviewDifferences,
  type ReviewOracleOutcome,
  storyRevisionIds,
} from "./reviewOracle";

setDefaultTimeout(propertyTestTimeout(240_000));

const resolved = (document: Document, decision: RevisionDecision): Document => {
  const result = applyDocumentOp(normalizeForOps(document), {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: storyRevisionIds(document),
    decision,
  });
  if (result.isErr()) throw result.error;
  return result.value.document;
};

type Finding = {
  operation: string;
  outcome: ReviewOracleOutcome;
  occurrences: number;
};
const reports = new Map<string, Finding>();
const counts = new Map<string, number>();
const recordOutcome = (
  operation: string,
  outcome: ReviewOracleOutcome,
): void => {
  const countKey = `${operation}:${outcome.type}`;
  counts.set(countKey, (counts.get(countKey) ?? 0) + 1);
  if (outcome.type === "match") return;
  const key = JSON.stringify({ operation, outcome });
  const previous = reports.get(key);
  reports.set(key, {
    operation,
    outcome,
    occurrences: (previous?.occurrences ?? 0) + 1,
  });
};
afterAll(() => {
  process.stdout.write(
    `${JSON.stringify({ oracle: "L9", counts: Object.fromEntries(counts), findings: [...reports.values()] })}\n`,
  );
});

/**
 * A tracked document reads back from the editor and from a saved file with
 * the same review state, and resolves the same way live as reopened.
 */
const expectSurvivesSaveAndReopen = async (pending: Document) => {
  expect(
    reviewDifferences(pending, editorProjectionRoundTrip(pending)),
  ).toEqual({ messages: [], omitted: 0 });
  const reopened = await editorRoundTrip(pending);
  expect(reviewDifferences(pending, reopened)).toEqual({
    messages: [],
    omitted: 0,
  });
  expect(storyRevisionIds(reopened)).toEqual(storyRevisionIds(pending));
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const liveResolved = resolved(pending, decision);
    const reopenedResolved = resolved(reopened, decision);
    expect(reviewDifferences(liveResolved, reopenedResolved)).toEqual({
      messages: [],
      omitted: 0,
    });
    expect(
      reviewDifferences(reopenedResolved, await editorRoundTrip(liveResolved)),
    ).toEqual({ messages: [], omitted: 0 });
  }
};

/** Whether a paragraph mark records which paragraph rejecting it keeps. */
const marksASurvivor = (document: Document): boolean =>
  JSON.stringify(document.package.document.content).includes('"survivor"');

describe("tracked model round trip and resolution oracle", () => {
  test("L8: an insertion after the paragraph ending a block list survives editor save and reopen", async () => {
    await assertProperty(
      fc.asyncProperty(seedArbitrary, async (seed) => {
        const document = await fixture(seed);
        const applied = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          story: OP_STORIES.MAIN,
          at: { type: "after", blockId: "00000004" },
          blocks: [
            {
              type: "paragraph",
              paraId: "00000021",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: seed.insertion }],
                },
              ],
            },
          ],
          revision: stamp,
          newIds,
        });
        if (applied.isErr()) throw applied.error;
        // The anchor ends its block list only inside the fixture's cell.
        fc.pre(marksASurvivor(applied.value.document));
        await expectSurvivesSaveAndReopen(applied.value.document);
      }),
      {
        numRuns: 20,
        id: "L8: an insertion after the paragraph ending a block list survives editor save and reopen",
      },
    );
  });

  for (const [kind, operation] of Object.entries(operations)) {
    test(`L8: ${kind} survives editor save and reopen`, async () => {
      await assertProperty(
        fc.asyncProperty(seedArbitrary, async (seed) => {
          const document = await fixture(seed);
          const applied = applyDocumentOp(
            document,
            operation({ document, seed }),
          );
          if (applied.isErr()) throw applied.error;
          expect(applied.value.revisions.length).toBeGreaterThan(0);
          await expectSurvivesSaveAndReopen(applied.value.document);
        }),
        { numRuns: 20, id: "L8: ${kind} survives editor save and reopen" },
      );
    });

    test(`L9: ${kind} reports model and editor resolution differences`, async () => {
      const law = (
        tally: { attempted: number },
        record: typeof recordOutcome,
      ) =>
        fc.asyncProperty(seedArbitrary, async (seed) => {
          const document = await fixture(seed);
          const applied = applyDocumentOp(
            document,
            operation({ document, seed }),
          );
          if (applied.isErr()) throw applied.error;
          for (const decision of Object.values(REVISION_DECISIONS)) {
            const outcome = compareReviewResolution(
              applied.value.document,
              decision,
            );
            switch (outcome.type) {
              case "match":
                record(kind, outcome);
                tally.attempted += 1;
                break;
              case "disagreement":
              case "refused":
              case "editor-failed":
                record(kind, outcome);
                tally.attempted += 1;
                break;
              case "invalid":
              case "no-revisions":
                panic("Generated oracle case did not exercise resolution", {
                  kind,
                  outcome,
                });
              default:
                outcome satisfies never;
            }
          }
        });
      await assertProperty(law({ attempted: 0 }, recordOutcome), {
        numRuns: 20,
        id: "L9: ${kind} reports model and editor resolution differences",
      });
      // A fixed sample reaches resolution; one case is enough. Its outcomes
      // stay out of the report, which counts the property's cases.
      const RUNS = 5;
      const tally = { attempted: 0 };
      await checkCoverageSample(
        law(tally, () => undefined),
        RUNS,
      );
      expect(tally.attempted).toBeGreaterThan(0);
    });
  }

  test("L8: deletion inside another insertion keeps both wrapper ids", async () => {
    await assertProperty(
      fc.asyncProperty(seedArbitrary, async (seed) => {
        const document = await fixture({ ...seed, nesting: "plain" });
        const insertion = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at: at(0),
          text: seed.insertion,
          runProps: INHERIT_RUN_PROPS,
          revision: { ...stamp, author: "Earlier" },
          newIds,
        });
        if (insertion.isErr()) throw insertion.error;
        const deletion = applyDocumentOp(insertion.value.document, {
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: at(0),
          to: at(seed.insertion.length),
          revision: {
            ...stamp,
            id: 2000,
            author: "Later",
            date: "2026-07-08T09:10:11.000Z",
          },
          newIds: {
            revision: Array.from({ length: 32 }, (_, index) => 2001 + index),
          },
        });
        if (deletion.isErr()) throw deletion.error;
        const pending = deletion.value.document;
        const target = findParagraph(
          pending.package.document.content,
          FIRST_ID,
        );
        expect(
          target?.content.some(
            (content) =>
              content.type === "insertion" &&
              content.content.some((child) => child.type === "deletion"),
          ),
        ).toBe(true);
        const reopened = await editorRoundTrip(pending);
        expect(reviewDifferences(pending, reopened)).toEqual({
          messages: [],
          omitted: 0,
        });
        expect(storyRevisionIds(reopened)).toEqual(storyRevisionIds(pending));
      }),
      {
        numRuns: 20,
        id: "L8: deletion inside another insertion keeps both wrapper ids",
      },
    );
  });
});
