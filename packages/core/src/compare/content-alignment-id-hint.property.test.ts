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
    "unique surviving text keeps its pairing when a deleted tail leaves its mark id",
    () => {
      assertProperty(
        fc.property(
          fc.record({
            surviving: fc.integer({ min: 1, max: 8 }),
            deleted: fc.integer({ min: 1, max: 4 }),
            inserted: fc.boolean(),
            idBase: fc.integer({ min: 0x10000000, max: 0x60000000 }),
            wording: fc.array(
              fc.constantFrom("alpha", "beta", "gamma", "delta"),
              {
                minLength: 8,
                maxLength: 12,
              },
            ),
          }),
          ({ surviving, deleted, inserted, idBase, wording }) => {
            for (const kind of ["paragraph", "heading", "listItem"]) {
              const scopes = [
                {},
                { containerPath: [{ kind: "control", id: "tail-control" }] },
                {
                  table: {
                    outerTableIndex: 0,
                    tableIndex: 0,
                    rowIndex: 0,
                    cellIndex: 0,
                    gridColumnIndex: 0,
                    columnSpan: 1,
                    rowSpan: 1,
                    paragraphIndex: 0,
                  },
                },
              ] satisfies Partial<FolioContentBlockIdentity>[];
              for (const scope of scopes) {
                const base = Array.from(
                  { length: surviving + deleted },
                  (_, index) => ({
                    id: (idBase + index).toString(16).toUpperCase(),
                    kind,
                    text: `${wording.join(" ")} clause ${String(index)}.`,
                    ...scope,
                  }),
                );
                const revised = base.slice(0, surviving).map((block, index) =>
                  Object.assign({}, block, {
                    // Deleting the final run leaves its paragraph mark on the
                    // preceding text; the former mark's id has been retired.
                    id:
                      index === surviving - 1
                        ? (idBase + surviving + deleted - 1)
                            .toString(16)
                            .toUpperCase()
                        : block.id,
                  }),
                );
                const addition = {
                  id: "70000000",
                  kind,
                  text: "An independently inserted paragraph.",
                  ...scope,
                };
                // Keep the insertion separated from the deleted tail by an
                // exact survivor; an adjacent deletion/insertion may be one edit.
                const insertionIndex = Math.min(1, surviving - 1);
                if (inserted) revised.splice(insertionIndex, 0, addition);
                const unchanged = base.slice(0, surviving).map(({ text }) => ({
                  pair: [text, text],
                }));
                const removed = base.slice(surviving);
                for (const exhausted of [false, true]) {
                  const session = () =>
                    exhausted
                      ? createFolioContentAlignmentWorkSession({
                          lcsCells: 0,
                          structuralTokenLookups: 0,
                        })
                      : createFolioContentAlignmentWorkSession();
                  expect(
                    eventsRead(
                      alignFolioContentBlocks(base, revised, {
                        workSession: session(),
                      }),
                    ),
                  ).toEqual([
                    ...unchanged.slice(0, insertionIndex),
                    ...(inserted ? [{ revisedOnly: addition.text }] : []),
                    ...unchanged.slice(insertionIndex),
                    ...removed.map(({ text }) => ({ baseOnly: text })),
                  ]);
                  expect(
                    eventsRead(
                      alignFolioContentBlocks(revised, base, {
                        workSession: session(),
                      }),
                    ),
                  ).toEqual([
                    ...unchanged.slice(0, insertionIndex),
                    ...(inserted ? [{ baseOnly: addition.text }] : []),
                    ...unchanged.slice(insertionIndex),
                    ...removed.map(({ text }) => ({ revisedOnly: text })),
                  ]);
                }
              }
            }
          },
        ),
        {
          id: "deleted tail mark identity preserves the unique surviving text",
          numRuns: 80,
          examples: [
            [
              {
                surviving: 2,
                deleted: 1,
                inserted: true,
                idBase: 0x10000000,
                wording: Array.from({ length: 8 }, () => "alpha"),
              },
            ],
          ],
        },
      );
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

  test("a unilateral counterpart must have an unambiguous retired or new stable mark in the same scope", () => {
    const source = {
      id: "10000001",
      kind: "paragraph",
      text: "The parties shall observe all obligations in this agreement alpha",
    } satisfies FolioContentBlockIdentity;
    const tail = {
      id: "10000002",
      kind: "paragraph",
      text: "The parties shall observe all obligations in this agreement beta",
    } satisfies FolioContentBlockIdentity;
    const survivor = { ...source, id: tail.id };
    const cell = {
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    };
    const cases = [
      {
        // A duplicate text cannot establish the survivor's former identity.
        base: [source, { ...source, id: "10000003" }, tail],
        revised: [survivor],
      },
      {
        // Nor can an exact source whose stable id belongs to two paragraphs.
        base: [source, { ...source, text: `${source.text} gamma` }, tail],
        revised: [survivor],
      },
      {
        base: [{ ...source, idStability: "positional" }, tail],
        revised: [survivor],
      },
      {
        // An id still present on a positional block is not a retired id.
        base: [source, tail],
        revised: [
          {
            ...source,
            idStability: "positional",
            text: `${source.text} gamma`,
          },
          survivor,
        ],
      },
      {
        // Duplicated evidence on the other side is not absence of evidence.
        base: [source, tail],
        revised: [
          survivor,
          { ...tail, id: "10000003" },
          { ...tail, id: "10000004" },
        ],
      },
      {
        base: [
          {
            ...source,
            containerPath: [{ kind: "control", id: "other-control" }],
          },
          tail,
        ],
        revised: [survivor],
      },
      {
        base: [
          { ...source, table: { ...cell, cellIndex: 1, gridColumnIndex: 1 } },
          { ...tail, table: cell },
        ],
        revised: [{ ...survivor, table: cell }],
      },
    ] satisfies {
      base: FolioContentBlockIdentity[];
      revised: FolioContentBlockIdentity[];
    }[];
    for (const { base, revised } of cases) {
      for (const exhausted of [false, true]) {
        for (const [before, after] of [
          [base, revised],
          [revised, base],
        ]) {
          const pairs = alignFolioContentBlocks(before, after, {
            workSession: createFolioContentAlignmentWorkSession(
              exhausted ? { lcsCells: 0, structuralTokenLookups: 0 } : {},
            ),
          }).flatMap((event) =>
            event.type === "pair"
              ? [[event.baseBlock.id, event.revisedBlock.id]]
              : [],
          );
          expect(pairs).toContainEqual([tail.id, tail.id]);
        }
      }
    }
    // With no similarity budget, blank text still supplies no exact evidence.
    for (const blank of ["", " ", "\t\n"]) {
      const events = alignFolioContentBlocks(
        [{ ...source, text: blank }, tail],
        [{ ...survivor, text: blank }],
        {
          workSession: createFolioContentAlignmentWorkSession({ lcsCells: 0 }),
        },
      );
      expect(events.find((event) => event.type === "pair")).toMatchObject({
        baseBlock: { id: tail.id },
        revisedBlock: { id: tail.id },
      });
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
