/**
 * A paragraph id is a hint for pairing. Unique exact counterparts override
 * contradictory ids even when paragraphs share most wording; ambiguous or
 * structurally incompatible text keeps the existing fuzzy matching policy.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  alignFolioContentBlocks,
  alignFolioContentStructure,
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

test(
  "a saved minted mark cannot outrank a more similar surviving paragraph",
  () => {
    fc.assert(
      fc.property(fc.integer({ min: 1, max: 8 }), (sharedWords) => {
        const shared = wordsOf(0, sharedWords);
        for (const baseStability of ["stable", "positional"] as const) {
          for (const revisedStability of ["stable", "positional"] as const) {
            const base = [
              {
                id: "10000000",
                kind: "paragraph",
                text: `Original ${shared}`,
                idStability: baseStability,
              },
              {
                id: "10000001",
                kind: "paragraph",
                text: `${wordsOf(1, sharedWords * 3)} ${shared}`,
                idStability: baseStability,
              },
            ] as const satisfies readonly FolioContentBlockIdentity[];
            const revised = [
              {
                id: "10000001",
                kind: "paragraph",
                text: `Edited ${shared}`,
                idStability: revisedStability,
              },
            ] as const satisfies readonly FolioContentBlockIdentity[];
            const expected = [
              { pair: [base[0].text, revised[0].text] },
              { baseOnly: base[1].text },
            ];
            expect(eventsRead(alignFolioContentBlocks(base, revised))).toEqual(
              expected,
            );
            expect(eventsRead(alignFolioContentBlocks(revised, base))).toEqual([
              { pair: [revised[0].text, base[0].text] },
              { revisedOnly: base[1].text },
            ]);
          }
        }
      }),
      propertyConfig({ numRuns: 40 }),
    );
  },
  propertyTestTimeout(60_000),
);

test("mixed identity continuity survives ambiguous, out-of-scope or unmeasured alternatives", () => {
  const source = {
    id: "10000001",
    kind: "paragraph",
    text: "retired longer wording paragraph",
    idStability: "positional",
  } as const;
  const target = {
    ...source,
    text: "edited paragraph",
    idStability: "stable",
  } as const;
  const stronger = { ...source, id: "10000000", text: "original paragraph" };
  const cases = [
    { alternatives: [{ ...stronger, text: "" }], budget: {} },
    { alternatives: [{ ...stronger, text: "!!!" }], budget: {} },
    { alternatives: [stronger, { ...stronger, id: "10000002" }], budget: {} },
    {
      alternatives: [
        { ...stronger, containerPath: [{ kind: "control", id: "different" }] },
      ],
      budget: {},
    },
    { alternatives: [stronger], budget: { lcsCells: 0 } },
    { alternatives: [stronger], budget: { structuralTokenLookups: 0 } },
  ] as const satisfies readonly {
    alternatives: readonly FolioContentBlockIdentity[];
    budget: { lcsCells?: number; structuralTokenLookups?: number };
  }[];
  for (const { alternatives, budget } of cases) {
    const base = [...alternatives, source];
    for (const [before, after] of [
      [base, [target]],
      [[target], base],
    ]) {
      if (before === undefined || after === undefined)
        panic("Both alignment sides exist");
      const pairs = alignFolioContentBlocks(before, after, {
        workSession: createFolioContentAlignmentWorkSession(budget),
      }).flatMap((event) =>
        event.type === "pair"
          ? [[event.baseBlock.id, event.revisedBlock.id]]
          : [],
      );
      expect(pairs).toContainEqual([source.id, target.id]);
    }
  }
});

test("a stronger candidate keeps the edit that owns its id beyond the current gap", () => {
  const competitor = {
    id: "10000000",
    kind: "paragraph",
    text: "original paragraph",
    idStability: "positional",
  } as const;
  const source = {
    ...competitor,
    id: "10000001",
    text: "different extended wording paragraph",
  };
  const anchor = { ...competitor, id: "10000002", text: "An exact separator." };
  const base = [competitor, source, anchor];
  const revised = [
    { ...source, text: "edited paragraph", idStability: "stable" as const },
    { ...anchor, idStability: "stable" as const },
    {
      ...competitor,
      text: "New distant content.",
      idStability: "stable" as const,
    },
  ];
  for (const [before, after] of [
    [base, revised],
    [revised, base],
  ]) {
    if (before === undefined || after === undefined)
      panic("Both alignment sides exist");
    const pairs = alignFolioContentBlocks(before, after).flatMap((event) =>
      event.type === "pair"
        ? [[event.baseBlock.id, event.revisedBlock.id]]
        : [],
    );
    expect(pairs).toContainEqual([source.id, source.id]);
  }
});

test("continuity with no retired candidates preserves the later gap allowance", () => {
  const block = ({ id, text }: { id: string; text: string }) =>
    ({
      id,
      text,
      kind: "paragraph",
      idStability: "positional",
    }) as const;
  const base = [
    block({ id: "10000000", text: "Earlier original wording" }),
    block({ id: "10000001", text: "An exact separator" }),
    block({ id: "10000002", text: "Unrelated inserted text" }),
    block({ id: "10000003", text: "Original surviving paragraph" }),
  ];
  const revised = [
    {
      ...block({ id: "10000000", text: "Earlier edited wording" }),
      idStability: "stable" as const,
    },
    block({ id: "10000001", text: "An exact separator" }),
    block({ id: "10000004", text: "Edited surviving paragraph" }),
  ];
  for (const [before, after] of [
    [base, revised],
    [revised, base],
  ]) {
    if (before === undefined || after === undefined) panic("Both sides exist");
    const pairs = alignFolioContentBlocks(before, after, {
      workSession: createFolioContentAlignmentWorkSession({ lcsCells: 3 }),
    }).flatMap((event) =>
      event.type === "pair"
        ? [[event.baseBlock.id, event.revisedBlock.id]]
        : [],
    );
    expect(pairs).toContainEqual(
      before === base ? ["10000003", "10000004"] : ["10000004", "10000003"],
    );
  }
});

test("incompatible retired candidates cannot consume a later gap's token allowance", () => {
  const source = {
    id: "10000001",
    kind: "paragraph",
    text: "red alpha",
    idStability: "positional",
  } as const;
  const anchor = {
    ...source,
    id: "10000003",
    text: "Exact separator",
    idStability: "stable",
  } as const;
  for (const irrelevant of [
    { ...source, id: "10000002", kind: "figure", text: "blue beta" },
    {
      ...source,
      id: "10000002",
      text: "blue beta",
      containerPath: [{ kind: "control", id: "other" }],
    },
  ] as const satisfies readonly FolioContentBlockIdentity[]) {
    const base = [
      source,
      irrelevant,
      anchor,
      { ...source, id: "10000004", text: "remote unrelated" },
      { ...source, id: "10000005", text: "original paragraph" },
    ];
    const revised = [
      { ...source, text: "green gamma", idStability: "stable" as const },
      anchor,
      { ...source, id: "10000006", text: "edited paragraph" },
    ];
    for (const [before, after] of [
      [base, revised],
      [revised, base],
    ]) {
      if (before === undefined || after === undefined)
        panic("Both sides exist");
      const pairs = alignFolioContentBlocks(before, after, {
        workSession: createFolioContentAlignmentWorkSession({
          lcsCells: 100,
          structuralTokenLookups: 4,
        }),
      }).flatMap((event) =>
        event.type === "pair"
          ? [[event.baseBlock.id, event.revisedBlock.id]]
          : [],
      );
      expect(pairs).toContainEqual(
        before === base ? ["10000005", "10000006"] : ["10000006", "10000005"],
      );
    }
  }
});

test("rounded similarity ties retain mixed identity regardless of candidate order and label", () => {
  const targetWords = Array.from(
    { length: 32 },
    (_, index) => `target${index}`,
  );
  const source = {
    id: "10000000",
    kind: "paragraph",
    idStability: "positional",
    text: [
      ...targetWords.slice(0, 20),
      ...Array.from({ length: 9 }, (_, index) => `source${index}`),
    ].join(" "),
  } as const;
  const candidate = {
    ...source,
    id: "10000001",
    displayLabel: "1.",
    text: [
      ...targetWords.slice(0, 21),
      ...Array.from({ length: 11 }, (_, index) => `candidate${index}`),
    ].join(" "),
  };
  const target = {
    ...source,
    idStability: "stable",
    text: targetWords.join(" "),
    displayLabel: "1.",
  } as const;
  for (const base of [
    [source, candidate],
    [candidate, source],
  ]) {
    for (const [before, after] of [
      [base, [target]],
      [[target], base],
    ]) {
      if (before === undefined || after === undefined)
        panic("Both sides exist");
      const pairs = alignFolioContentBlocks(before, after).flatMap((event) =>
        event.type === "pair"
          ? [[event.baseBlock.id, event.revisedBlock.id]]
          : [],
      );
      expect(pairs).toContainEqual([source.id, target.id]);
    }
  }
});

test("retired identity evidence uses the complete document across structural partitions", () => {
  const table = (cellIndex: number, rowIndex = 0, tableIndex = 0) => ({
    outerTableIndex: tableIndex,
    tableIndex,
    rowIndex,
    cellIndex,
    gridColumnIndex: cellIndex,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex: 0,
  });
  const competitor = {
    id: "10000000",
    kind: "paragraph",
    text: "original paragraph",
    idStability: "positional",
    table: table(0),
  } as const;
  const source = {
    ...competitor,
    id: "10000001",
    text: "different extended wording paragraph",
  };
  const anchor = {
    ...competitor,
    id: "10000002",
    text: "Exact cell separator",
  };
  for (const elsewhere of [table(1), table(0, 1), table(0, 0, 1), undefined]) {
    const other = {
      ...competitor,
      id: "10000003",
      text: "Exact distant separator",
      table: elsewhere,
    };
    const base = [competitor, source, anchor, other];
    const revised = [
      { ...source, text: "edited paragraph", idStability: "stable" as const },
      { ...anchor, idStability: "stable" as const },
      { ...other, idStability: "stable" as const },
      {
        ...competitor,
        table: elsewhere,
        text: "New distant content",
        idStability: "stable" as const,
      },
    ];
    for (const [before, after] of [
      [base, revised],
      [revised, base],
    ]) {
      if (before === undefined || after === undefined)
        panic("Both sides exist");
      const pairs = alignFolioContentStructure({
        baseBlocks: before,
        revisedBlocks: after,
      }).flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.id, step.revisedBlock.id]] : [],
      );
      expect(pairs).toContainEqual([source.id, source.id]);
    }
  }
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
