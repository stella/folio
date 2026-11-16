/**
 * A paragraph id is a hint for pairing. Unique exact counterparts override
 * contradictory ids even when paragraphs share most wording; ambiguous or
 * structurally incompatible text keeps the existing fuzzy matching policy.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  alignFolioContentBlocks,
  createFolioContentAlignmentWorkSession,
  type FolioContentAlignedBlockEvent,
} from "./content-alignment";
import type { FolioContentBlockIdentity } from "./content-types";

/** Words no two paragraphs share: paragraph `index` uses its own. */
const wordsOf = (index: number, count: number): string =>
  Array.from(
    { length: count },
    (_, word) => `w${String(index)}x${String(word)}`,
  ).join(" ");

const storyArbitrary = fc
  .array(fc.integer({ min: 1, max: 6 }), { minLength: 2, maxLength: 12 })
  .chain((lengths) =>
    fc.record({
      lengths: fc.constant(lengths),
      // Which base paragraph's id each revised paragraph carries.
      order: fc.shuffledSubarray(
        lengths.map((_, index) => index),
        { minLength: lengths.length, maxLength: lengths.length },
      ),
    }),
  );

/** Paragraph word counts, and which base paragraph's id each revised one carries. */
type Story = { lengths: readonly number[]; order: readonly number[] };

const blocksOf = (
  story: Story,
  kind: string,
): {
  base: FolioContentBlockIdentity[];
  revised: FolioContentBlockIdentity[];
} => {
  const texts = story.lengths.map((length, index) => wordsOf(index, length));
  const idOf = (index: number) =>
    (0x10000000 + index).toString(16).toUpperCase();
  return {
    base: texts.map((text, index) => ({ id: idOf(index), kind, text })),
    revised: texts.map((text, index) => ({
      id: idOf(story.order[index] ?? index),
      kind,
      text,
    })),
  };
};

/** Each event as the texts it pairs, or the side that holds it alone. */
const eventsRead = (events: readonly FolioContentAlignedBlockEvent[]) =>
  events.map((event) => {
    switch (event.type) {
      case "pair":
        return { pair: [event.baseBlock.text, event.revisedBlock.text] };
      case "baseOnly":
        return { baseOnly: event.block.text };
      case "revisedOnly":
        return { revisedOnly: event.block.text };
      default: {
        const unreachable: never = event;
        return unreachable;
      }
    }
  });

const law = (kind: string) =>
  fc.property(storyArbitrary, (story) => {
    const { base, revised } = blocksOf(story, kind);
    expect(eventsRead(alignFolioContentBlocks(base, revised))).toEqual(
      base.map(({ text }) => ({ pair: [text, text] })),
    );
  });

/** Every paragraph keeps enough shared wording to satisfy the fuzzy pairing threshold. */
const similarBlocksOf = (story: Story, kind: string) => {
  const prepared = blocksOf(story, kind);
  const shared = "The parties shall observe all obligations in this agreement";
  const prefix = (block: FolioContentBlockIdentity) =>
    Object.assign({}, block, { text: `${shared} ${block.text}` });
  return {
    base: prepared.base.map(prefix),
    revised: prepared.revised.map(prefix),
  };
};

describe("a paragraph id the content does not agree with", () => {
  test(
    "pairs a story with itself as if the ids were not there",
    () => {
      assertProperty(law("paragraph"), {
        id: "shuffled paragraph ids pair a story with itself",
        numRuns: 200,
      });
    },
    propertyTestTimeout(60_000),
  );

  test(
    "unique exact counterparts outweigh rebound ids despite similar wording",
    () => {
      assertProperty(
        fc.property(storyArbitrary, (story) => {
          for (const kind of ["paragraph", "heading", "listItem"]) {
            const { base, revised } = similarBlocksOf(story, kind);
            for (const exhausted of [false, true]) {
              const events = alignFolioContentBlocks(base, revised, {
                workSession: exhausted
                  ? createFolioContentAlignmentWorkSession({
                      lcsCells: 0,
                      structuralTokenLookups: 0,
                    })
                  : createFolioContentAlignmentWorkSession(),
              });
              expect(eventsRead(events)).toEqual(
                base.map(({ text }) => ({ pair: [text, text] })),
              );
            }
          }
        }),
        {
          id: "unique exact paragraph counterparts outweigh rebound stable ids",
          numRuns: 150,
          examples: [[{ lengths: [2, 3, 4], order: [2, 1, 0] }]],
        },
      );
    },
    propertyTestTimeout(60_000),
  );

  test("one-sided, duplicate or cross-container text does not release similar stable ids", () => {
    const text = (word: string) =>
      `The parties shall observe all obligations in this agreement ${word}`;
    const paragraph = (
      id: string,
      word: string,
    ): FolioContentBlockIdentity => ({
      id,
      kind: "paragraph",
      text: text(word),
    });
    const cases = [
      {
        base: [paragraph("A1000001", "alpha"), paragraph("A1000002", "gamma")],
        revised: [
          paragraph("A1000001", "beta"),
          paragraph("A1000002", "alpha"),
        ],
      },
      {
        base: [
          paragraph("A1000001", "alpha"),
          paragraph("A1000002", "beta"),
          paragraph("A1000003", "alpha"),
        ],
        revised: [
          paragraph("A1000001", "beta"),
          paragraph("A1000002", "alpha"),
        ],
      },
      {
        base: [paragraph("A1000001", "alpha"), paragraph("A1000002", "beta")],
        revised: [
          paragraph("A1000001", "beta"),
          paragraph("A1000002", "alpha"),
          paragraph("A1000003", "beta"),
        ],
      },
      {
        base: [
          paragraph("A1000001", "alpha"),
          Object.assign(paragraph("A1000002", "beta"), {
            containerPath: [{ kind: "control", id: "container-one" }],
          }),
        ],
        revised: [
          paragraph("A1000001", "beta"),
          Object.assign(paragraph("A1000002", "alpha"), {
            containerPath: [{ kind: "control", id: "container-one" }],
          }),
        ],
      },
    ];
    for (const { base, revised } of cases) {
      const pairs = alignFolioContentBlocks(base, revised).flatMap((event) =>
        event.type === "pair"
          ? [[event.baseBlock.id, event.revisedBlock.id]]
          : [],
      );
      expect(pairs).toContainEqual(["A1000001", "A1000001"]);
      expect(pairs).toContainEqual(["A1000002", "A1000002"]);
    }
  });

  test(
    "the law catches an id trusted without its content",
    () => {
      // Blocks outside the paragraph family keep their id's word, so the
      // same shuffle pairs them by id across their content.
      const result = fc.check(
        law("drawing"),
        propertyConfig({
          numRuns: 1,
          examples: [[{ lengths: [2, 3, 4], order: [2, 1, 0] }]],
          reporter: () => undefined,
        }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(60_000),
  );
});
