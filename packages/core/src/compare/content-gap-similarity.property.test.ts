import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { compareContent, type FolioContentComparisonEvent } from "./content";
import type { FolioContentBlock } from "./content-types";

setDefaultTimeout(propertyTestTimeout(30_000));

// Every block speaks its own vocabulary, so no two blocks are similar unless
// one is the other with words added.
const blockText = (block: number): string =>
  Array.from({ length: 6 }, (_, word) => `b${String(block)}w${String(word)}`).join(" ");

const APPENDED = "appended words close the paragraph.";

const toBlocks = (texts: readonly string[], side: string): FolioContentBlock[] =>
  texts.map((text, index) => ({
    id: `${side}-${String(index)}`,
    idStability: "positional",
    kind: "paragraph",
    text,
  }));

type Edit = {
  /** Distinct original blocks. */
  count: number;
  /** Where each new block goes, as an index into the growing list. */
  insertions: readonly number[];
  /** The original block that gains `APPENDED`. */
  edited: number;
};

const editArbitrary = fc.integer({ min: 1, max: 8 }).chain((count) =>
  fc.record({
    count: fc.constant(count),
    insertions: fc.array(fc.nat(), { minLength: 1, maxLength: 4 }),
    edited: fc.integer({ min: 0, max: count - 1 }),
  }),
);

type Revision = { base: FolioContentBlock[]; revised: FolioContentBlock[] };

/** The original list, and the list with new blocks inserted and one block extended. */
const revise = ({ count, edited, insertions }: Edit): Revision => {
  const original = Array.from({ length: count }, (_, block) => blockText(block));
  const revised = original.map((text, block) => (block === edited ? `${text} ${APPENDED}` : text));
  for (const [ordinal, position] of insertions.entries()) {
    revised.splice(position % (revised.length + 1), 0, blockText(count + ordinal));
  }
  return { base: toBlocks(original, "base"), revised: toBlocks(revised, "revised") };
};

const events = (base: FolioContentBlock[], revised: FolioContentBlock[]) =>
  compareContent({ base: { blocks: base }, revised: { blocks: revised } }).unwrap().events;

const countOf = (list: readonly FolioContentComparisonEvent[], type: string): number =>
  list.filter((event) => event.type === type).length;

const changedText = (event: FolioContentComparisonEvent, type: "ins" | "del"): string =>
  event.type === "modified"
    ? event.segments
        .filter((segment) => segment.type === type)
        .map((segment) => segment.text)
        .join("")
        .trim()
    : "";

describe("content comparison inside a gap that gained or lost blocks", () => {
  test("new blocks read as inserted and the edited neighbour as its own modification", () => {
    fc.assert(
      fc.property(editArbitrary, (edit) => {
        const { base, revised } = revise(edit);
        const compared = events(base, revised);

        expect(countOf(compared, "inserted")).toBe(edit.insertions.length);
        expect(countOf(compared, "deleted")).toBe(0);
        expect(countOf(compared, "unchanged")).toBe(edit.count - 1);
        const modified = compared.filter((event) => event.type === "modified");
        expect(modified).toHaveLength(1);
        const [only] = modified;
        expect(only?.baseBlocks[0].text).toBe(blockText(edit.edited));
        expect(only && changedText(only, "ins")).toBe(APPENDED);
        expect(only && changedText(only, "del")).toBe("");
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("swapping the sides swaps insertions for deletions and keeps the pairing", () => {
    fc.assert(
      fc.property(editArbitrary, (edit) => {
        const { base, revised } = revise(edit);
        const forward = events(base, revised);
        const backward = events(revised, base);

        expect(countOf(backward, "deleted")).toBe(countOf(forward, "inserted"));
        expect(countOf(backward, "inserted")).toBe(countOf(forward, "deleted"));
        const pairs = (list: readonly FolioContentComparisonEvent[], flip: boolean) =>
          list
            .filter((event) => event.type === "modified" || event.type === "unchanged")
            .map((event) => {
              const [older, newer] = [event.baseBlocks[0]?.text, event.revisedBlocks[0]?.text];
              return flip
                ? `${String(newer)}|${String(older)}`
                : `${String(older)}|${String(newer)}`;
            });
        expect(pairs(backward, true)).toEqual(pairs(forward, false));
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a paragraph split or merged at any word reads as one split or merge", () => {
    fc.assert(
      fc.property(fc.integer({ min: 2, max: 12 }), fc.nat(), (length, cut) => {
        const words = Array.from({ length }, (_, word) => `w${String(word)}`);
        const at = 1 + (cut % (length - 1));
        const whole = toBlocks([blockText(90), words.join(" "), blockText(91)], "whole");
        const halves = toBlocks(
          [blockText(90), words.slice(0, at).join(" "), words.slice(at).join(" "), blockText(91)],
          "halves",
        );

        expect(countOf(events(whole, halves), "split")).toBe(1);
        expect(countOf(events(halves, whole), "merge")).toBe(1);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("a block rewritten beyond recognition beside an insertion is not paired", () => {
    const base = toBlocks([blockText(0), blockText(1), blockText(2)], "base");
    const revised = toBlocks([blockText(0), blockText(3), blockText(4), blockText(2)], "revised");

    const compared = events(base, revised);
    expect(compared.map(({ type }) => type)).toEqual([
      "unchanged",
      "deleted",
      "inserted",
      "inserted",
      "unchanged",
    ]);
  });
});
