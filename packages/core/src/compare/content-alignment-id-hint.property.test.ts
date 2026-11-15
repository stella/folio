/**
 * A paragraph id is a hint for pairing, never proof. Comparing a story with
 * itself after its paragraph ids were shuffled among paragraphs that share no
 * wording pairs every paragraph with its own copy, as if the ids were not
 * there: an id is trusted only where the content agrees with it.
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
