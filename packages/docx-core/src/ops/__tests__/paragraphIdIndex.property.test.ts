/**
 * The seed contract keeps every paragraph id of a package unique, and an
 * operation is refused rather than leave one repeated.
 *
 * The package's paragraph id index, found from the package indexed before
 * it, counts what a census of every paragraph counts, and says how many ids
 * repeat. Runs of generated edits index every version, and versions with a
 * paragraph repeated by hand next to them, which the contract refuses.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { meetsContract, validateOpsDocument } from "../contract";
import { countIds, packageParagraphIdIndex, packageParagraphIds } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { documentStories } from "../stories";
import { DOCUMENT_OP_TYPES, INHERIT_RUN_PROPS, OP_STORIES } from "../types";
import {
  documentArbitrary,
  opForStory,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

/** The property's run count, which the threshold was set for. */
const COVERAGE_RUNS = 120;

const withBody = (
  document: Document,
  content: readonly BlockContent[],
): Document => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content: [...content] },
  },
});

const bodyParagraphs = (document: Document): Paragraph[] =>
  document.package.document.content.filter(
    (block): block is Paragraph => block.type === "paragraph",
  );

/**
 * The body with a copy of one of its paragraphs added after the last block,
 * carrying the same id, spelled in the other case when `recased`.
 */
const withRepeatedParagraph = (
  document: Document,
  pick: number,
  recased: boolean,
): Document | undefined => {
  const paragraphs = bodyParagraphs(document);
  const original = paragraphs[pick % Math.max(paragraphs.length, 1)];
  if (original?.paraId === undefined) return undefined;
  const lower = original.paraId.toLowerCase();
  const otherCase =
    lower === original.paraId ? original.paraId.toUpperCase() : lower;
  const paraId = recased ? otherCase : original.paraId;
  const content = document.package.document.content;
  return withBody(document, [...content, { ...original, paraId }]);
};

/** What a census of every paragraph counts, and how many ids it counts twice or more. */
const countedInFull = (document: Document) => {
  const counts = countIds(packageParagraphIds(document.package));
  return {
    counts,
    repeated: [...counts.values()].filter((count) => count > 1).length,
  };
};

const expectIndexed = (document: Document, listed: boolean): void => {
  const expected = countedInFull(document);
  const index = packageParagraphIdIndex(document.package);
  for (const [key, count] of expected.counts) {
    expect([key, index.count(key)]).toEqual([key, count]);
    expect(index.has(key)).toBe(true);
  }
  expect(index.has("7FFFFFFF")).toBe(expected.counts.has("7FFFFFFF"));
  expect(index.repeated()).toBe(expected.repeated);
  if (listed) expect(new Map(index.counts())).toEqual(expected.counts);
};

describe("the paragraph id index of a package", () => {
  test("equals a census of every paragraph, repeated ids included, over a run of edits", async () => {
    const law = (counts: { repeatedSeen: number }) =>
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.array(
          fc.record({
            seed: opSeedArbitrary,
            tracked: fc.boolean(),
            story: fc.nat(),
            // Edit an earlier version instead of the latest one.
            branch: fc.option(fc.nat(), { freq: 4 }),
            // A version with a paragraph repeated, next to the edited one.
            repeat: fc.option(
              fc.record({ pick: fc.nat(), recased: fc.boolean() }),
              { freq: 2 },
            ),
            listed: fc.boolean(),
          }),
          { minLength: 1, maxLength: 30 },
        ),
        (start, edits) => {
          const versions: Document[] = [start];
          expectIndexed(start, false);
          edits.forEach((edit, index) => {
            const latest = versions.at(-1) ?? start;
            const document =
              edit.branch === null
                ? latest
                : (versions[edit.branch % versions.length] ?? latest);
            if (edit.repeat !== null) {
              const repeated = withRepeatedParagraph(
                document,
                edit.repeat.pick,
                edit.repeat.recased,
              );
              if (repeated !== undefined) {
                expectIndexed(repeated, edit.listed);
                expect(
                  packageParagraphIdIndex(repeated.package).repeated(),
                ).toBe(1);
                const refused = validateOpsDocument(repeated);
                expect(refused.isErr() && refused.error.reason).toBe(
                  DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID,
                );
                counts.repeatedSeen += 1;
              }
            }
            const stories = documentStories(document);
            const story =
              stories[edit.story % stories.length] ?? OP_STORIES.MAIN;
            const op =
              edit.tracked && story === OP_STORIES.MAIN
                ? trackedOpFor(document, edit.seed, index)
                : opForStory({ document, seed: edit.seed, story });
            const applied = applyDocumentOp(document, op);
            if (applied.isErr()) return;
            const next = applied.value.document;
            expectIndexed(next, edit.listed);
            expect(packageParagraphIdIndex(next.package).repeated()).toBe(0);
            versions.push(next);
          });
          // Every version again, after later ones were found from them.
          for (const version of versions) expectIndexed(version, true);
        },
      );
    await assertProperty(law({ repeatedSeen: 0 }), {
      id: "paragraph id index follows edits soundly",
      numRuns: 120,
    });
    const counts = { repeatedSeen: 0 };
    checkCoverageSample(law(counts), COVERAGE_RUNS);
    expect(counts.repeatedSeen).toBeGreaterThan(100);
  });
});

describe("an edit that would leave a paragraph id repeated", () => {
  test("is refused, though the document it starts from was taken as meeting the contract", () => {
    const paragraph = (paraId: string, text: string): Paragraph => ({
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    });
    const start: Document = {
      package: {
        document: {
          content: [paragraph("1A2B3C4D", "one"), paragraph("2B3C4D5E", "two")],
        },
      },
    };
    expect(validateOpsDocument(start).isOk()).toBe(true);
    const op = {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: OP_STORIES.MAIN, blockId: "2B3C4D5E", offset: 0 },
      text: "x",
      runProps: INHERIT_RUN_PROPS,
    } as const;
    expect(applyDocumentOp(start, op).isOk()).toBe(true);
    // A document an operation would make, recorded as meeting the contract
    // without a check: the repeat it holds is not read again on input.
    const repeating = withBody(start, [
      ...start.package.document.content,
      paragraph("1a2b3c4d", "again"),
    ]);
    meetsContract(repeating);
    expect(validateOpsDocument(repeating).isOk()).toBe(true);
    const applied = applyDocumentOp(repeating, op);
    expect(applied.isErr() && applied.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID,
    );
  });
});

describe("a package indexed with no earlier package alive", () => {
  // Every package indexed before is collected, so the index counts this one
  // in full instead of finding it from another.
  test("counts its repeated ids", () => {
    const repeating = (ids: readonly string[]): Document => ({
      package: {
        document: {
          content: ids.map((paraId) => ({
            type: "paragraph" as const,
            paraId,
            content: [],
          })),
        },
      },
    });
    for (const ids of [
      ["11111111", "22222222"],
      ["11111111", "11111111", "22222222"],
      ["0000ABCD", "0000abcd", "22222222", "22222222", "22222222"],
    ]) {
      Bun.gc(true);
      const document = repeating(ids);
      expectIndexed(document, true);
      const refused = validateOpsDocument(document);
      const repeats = countedInFull(document).repeated;
      expect(refused.isErr() && refused.error.reason).toBe(
        repeats === 0 ? false : DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID,
      );
    }
  });
});
