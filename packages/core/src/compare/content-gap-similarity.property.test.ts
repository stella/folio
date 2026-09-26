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

const LABEL_MODES = ["text", "displayLabel", "none"] as const;
type LabelMode = (typeof LABEL_MODES)[number];

/** Two words every item shares, so unrelated items stay below the pairing threshold. */
const itemText = (tag: string): string =>
  `shall apply ${Array.from({ length: 5 }, (_, word) => `${tag}w${String(word)}`).join(" ")}`;
const REWORDED = " and further";

type Item = { origin: string; text: string; label?: string };
type Section = { scope: string; heading: string; items: readonly Item[] };

/** Ids carry the section and the item's origin, so pairings read side-independently. */
const sectionBlocks = (
  side: string,
  sections: readonly Section[],
  mode: LabelMode,
): FolioContentBlock[] =>
  sections.flatMap(({ scope, heading, items }): FolioContentBlock[] => [
    {
      id: `${side}|${scope}|heading`,
      idStability: "positional",
      kind: "heading",
      headingLevel: 2,
      text: heading,
    },
    ...items.map(({ origin, text, label: authored }, index): FolioContentBlock => {
      const label = authored ?? `${String.fromCodePoint(97 + index)})`;
      const block = {
        id: `${side}|${scope}|${origin}`,
        idStability: "positional",
        kind: "paragraph",
      } as const;
      switch (mode) {
        case "text":
          return { ...block, text: `${label} ${text}` };
        case "displayLabel":
          return { ...block, text, displayLabel: label };
        case "none":
          return { ...block, text };
        default: {
          const unreachable: never = mode;
          return unreachable;
        }
      }
    }),
  ]);

const originOf = (block: FolioContentBlock): string => block.id.slice(block.id.indexOf("|") + 1);

/** Every pairing a comparison made, as sorted `base -> revised` origins; `-` marks no partner. */
const pairingsOf = (list: readonly FolioContentComparisonEvent[]): string[] => {
  const movedFrom = new Map<string, FolioContentBlock>();
  for (const event of list) {
    if (event.type === "movedFrom") {
      for (const block of event.baseBlocks) {
        movedFrom.set(block.id, block);
      }
    }
  }
  return list
    .flatMap((event) => {
      switch (event.type) {
        case "unchanged":
        case "modified":
        case "formatting":
        case "split":
        case "merge":
          return [
            `${event.baseBlocks.map(originOf).join("+")} -> ${event.revisedBlocks.map(originOf).join("+")}`,
          ];
        case "inserted":
          return event.revisedBlocks.map((block) => `- -> ${originOf(block)}`);
        case "deleted":
          return event.baseBlocks.map((block) => `${originOf(block)} -> -`);
        case "movedFrom":
          return [];
        case "movedTo": {
          const from = movedFrom.get(event.baseBlockId);
          return event.revisedBlocks.map(
            (block) => `${from ? originOf(from) : "-"} -> ${originOf(block)}`,
          );
        }
        default: {
          const unreachable: never = event;
          return unreachable;
        }
      }
    })
    .toSorted();
};

const compareSections = (
  older: readonly Section[],
  newer: readonly Section[],
  mode: LabelMode,
): string[] =>
  pairingsOf(events(sectionBlocks("base", older, mode), sectionBlocks("revised", newer, mode)));

type Amendment = { older: readonly Item[]; newer: readonly Item[] };

/** A list with items deleted, inserted and at most one reworded; equal lengths included. */
const amendmentArbitrary = fc
  .record({
    count: fc.integer({ min: 1, max: 5 }),
    deletions: fc.array(fc.nat(), { maxLength: 2 }),
    insertions: fc.array(fc.nat(), { maxLength: 3 }),
    reworded: fc.option(fc.nat(), { nil: undefined }),
  })
  .filter(({ deletions, insertions }) => deletions.length + insertions.length > 0)
  .map(({ count, deletions, insertions, reworded }): Amendment => {
    const older = Array.from({ length: count }, (_, index) => ({
      origin: `old${String(index)}`,
      text: itemText(`o${String(index)}`),
    }));
    const newer = [...older];
    for (const position of deletions) {
      if (newer.length > 0) {
        newer.splice(position % newer.length, 1);
      }
    }
    const rewordedItem = reworded === undefined ? undefined : newer.at(reworded % newer.length);
    if (reworded !== undefined && rewordedItem) {
      newer[reworded % newer.length] = {
        origin: rewordedItem.origin,
        text: `${rewordedItem.text}${REWORDED}`,
      };
    }
    for (const [ordinal, position] of insertions.entries()) {
      newer.splice(position % (newer.length + 1), 0, {
        origin: `new${String(ordinal)}`,
        text: itemText(`n${String(ordinal)}`),
      });
    }
    return { older, newer };
  });

type Neighbour = { before: boolean; repeats: readonly (number | null)[] };

/** Unchanged sections whose items mostly repeat wording from either side of the amended list. */
const neighboursArbitrary = fc.array(
  fc.record({
    before: fc.boolean(),
    repeats: fc.array(fc.option(fc.nat(), { nil: null, freq: 4 }), {
      minLength: 1,
      maxLength: 3,
    }),
  }),
  { minLength: 1, maxLength: 3 },
);

const AMENDED = "amended";

const amendedPairings = (
  { older, newer }: Amendment,
  neighbours: readonly Neighbour[],
  mode: LabelMode,
): string[] => {
  const pool = [...older, ...newer];
  const unchanged = neighbours.map(({ before, repeats }, section) => ({
    before,
    section: {
      scope: `neighbour${String(section)}`,
      heading: `Section ${String(10 + section)}`,
      items: repeats.map((repeat, index) => ({
        origin: `item${String(index)}`,
        text:
          repeat === null
            ? itemText(`x${String(section)}i${String(index)}`)
            : (pool.at(repeat % pool.length)?.text ?? ""),
      })),
    },
  }));
  const around = (items: readonly Item[]): Section[] => [
    ...unchanged.filter(({ before }) => before).map(({ section }) => section),
    { scope: AMENDED, heading: "Section 5", items },
    ...unchanged.filter(({ before }) => !before).map(({ section }) => section),
  ];
  return compareSections(around(older), around(newer), mode).filter((pairing) =>
    pairing.includes(`${AMENDED}|`),
  );
};

describe("content comparison pairing inside any gap", () => {
  for (const mode of LABEL_MODES) {
    test(`neighbouring sections repeating the amended section's wording never change pairings inside it (labels: ${mode})`, () => {
      fc.assert(
        fc.property(amendmentArbitrary, neighboursArbitrary, (amendment, neighbours) => {
          expect(amendedPairings(amendment, neighbours, mode)).toEqual(
            amendedPairings(amendment, [], mode),
          );
        }),
        propertyConfig({ numRuns: 200 }),
      );
    });

    test(`every kept block pairs with itself, including equal-size gaps (labels: ${mode})`, () => {
      fc.assert(
        fc.property(amendmentArbitrary, neighboursArbitrary, (amendment, neighbours) => {
          const kept = amendment.older.filter(({ origin }) =>
            amendment.newer.some((item) => item.origin === origin),
          );
          for (const context of [[], neighbours]) {
            const pairings = amendedPairings(amendment, context, mode);
            for (const { origin } of kept) {
              expect(pairings).toContain(`${AMENDED}|${origin} -> ${AMENDED}|${origin}`);
            }
          }
        }),
        propertyConfig({ numRuns: 200 }),
      );
    });
  }

  const [A, B, C] = ["a", "b", "c"].map((tag) => ({ origin: tag, text: itemText(tag) }));
  const section = (scope: string, items: readonly (Item | undefined)[]): Section => ({
    scope,
    heading: `Section ${scope}`,
    items: items.filter((item) => item !== undefined),
  });

  test("an item inserted ahead of a list whose wording recurs later reads as one insertion", () => {
    expect(
      compareSections(
        [section("1", [A, B]), section("2", [A])],
        [section("1", [C, A, B]), section("2", [A])],
        "text",
      ),
    ).toEqual([
      "- -> 1|c",
      "1|a -> 1|a",
      "1|b -> 1|b",
      "1|heading -> 1|heading",
      "2|a -> 2|a",
      "2|heading -> 2|heading",
    ]);
  });

  test("an insertion beside a deletion in an equal gap keeps the surviving item", () => {
    expect(
      compareSections(
        [section("1", [A, B]), section("2", [A])],
        [section("1", [C, A]), section("2", [A])],
        "none",
      ),
    ).toEqual([
      "- -> 1|c",
      "1|a -> 1|a",
      "1|b -> -",
      "1|heading -> 1|heading",
      "2|a -> 2|a",
      "2|heading -> 2|heading",
    ]);
  });

  test("a reworded item beside a replaced one pairs with its own rewording", () => {
    const rewordedB = B && { origin: B.origin, text: `${B.text}${REWORDED}` };
    expect(compareSections([section("1", [A, B])], [section("1", [rewordedB, C])], "none")).toEqual(
      ["- -> 1|c", "1|a -> -", "1|b -> 1|b", "1|heading -> 1|heading"],
    );
  });

  test("renumbered duplicates keep their order around an insertion and a deletion", () => {
    const repeated = itemText("r");
    const [first, second] = ["r1", "r2"].map((origin) => ({ origin, text: repeated }));
    expect(
      compareSections(
        [section("1", [first, second, A])],
        [section("1", [C, first, second])],
        "displayLabel",
      ),
    ).toEqual(["- -> 1|c", "1|a -> -", "1|heading -> 1|heading", "1|r1 -> 1|r1", "1|r2 -> 1|r2"]);
  });

  test("identical wording pairs by its display label", () => {
    const omitted = (origin: string): Item => ({
      origin,
      text: "Intentionally omitted.",
      label: `(${origin})`,
    });
    expect(
      compareSections(
        [section("1", [omitted("1"), omitted("2"), A])],
        [section("1", [omitted("2"), A])],
        "displayLabel",
      ),
    ).toEqual(["1|1 -> -", "1|2 -> 1|2", "1|a -> 1|a", "1|heading -> 1|heading"]);
  });
});
