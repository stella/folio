import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { envelopes } from "../__tests__/wireFixtures";
import { applyDocumentOps } from "../apply";
import { equalForStaleness } from "../equality";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../types";
import { applyBatch } from "./applyBatch";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
} from "./envelope";
import { createSequencer } from "./sequencer";
import { transformBatch } from "./transform";

const at = (offset: number, blockId = "00000001") =>
  ({ story: "main", blockId, offset }) as const;
const insertion = (offset: number, text: string) =>
  ({
    type: "insertText",
    at: at(offset),
    text,
    runProps: "inherit",
  }) as const satisfies DocumentOp;
const deletion = (from: number, to: number) =>
  ({
    type: "deleteRange",
    from: at(from),
    to: at(to),
  }) as const satisfies DocumentOp;
const batch = (...ops: DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "incoming",
  actor: "actor",
  baseRev: 0,
  ops,
});
const sequenced = (...ops: DocumentOp[]): SequencedBatch => ({
  ...batch(...ops),
  opId: "foreign",
  revision: 1,
});
const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const documentOf = (text = "abcdef"): Document => ({
  package: { document: { content: [paragraph("00000001", text)] } },
});
const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value.document;
};
const transformed = (
  incoming: DocumentBatch,
  foreign: SequencedBatch,
  order: "before" | "after" = "after",
) => {
  const result = transformBatch(incoming, [foreign], { order });
  if (result.isErr()) throw result.error;
  return result.value.ops;
};
const converge = (
  document: Document,
  first: DocumentBatch,
  second: DocumentBatch,
) => {
  const firstSequenced = { ...first, revision: 1 };
  const secondSequenced = { ...second, revision: 1 };
  const left = apply(
    apply(document, first.ops),
    transformed(second, firstSequenced),
  );
  const right = apply(
    apply(document, second.ops),
    transformed(first, secondSequenced, "before"),
  );
  expect(equalForStaleness(left, right)).toBe(true);
};

describe("position transforms", () => {
  test(
    "insertions converge for every offset and UTF-16 width",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 6 }),
          fc.integer({ min: 0, max: 6 }),
          fc.constantFrom("X", "😀", "XY"),
          (left, right, text) => {
            converge(
              documentOf(),
              batch(insertion(left, text)),
              batch(insertion(right, "Z")),
            );
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "direct deletions preserve concurrent insertions at every boundary",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 6 }),
          fc.integer({ min: 1, max: 6 }),
          fc.integer({ min: 0, max: 6 }),
          (startSeed, widthSeed, point) => {
            const from = startSeed % 6;
            const to = from + 1 + (widthSeed % (6 - from));
            converge(
              documentOf(),
              batch(deletion(from, to)),
              batch(insertion(point, "X")),
            );
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "overlapping direct deletions converge",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 5 }),
          fc.integer({ min: 1, max: 6 }),
          fc.integer({ min: 0, max: 5 }),
          fc.integer({ min: 1, max: 6 }),
          (a, widthA, b, widthB) => {
            converge(
              documentOf(),
              batch(deletion(a, a + 1 + (widthA % (6 - a)))),
              batch(deletion(b, b + 1 + (widthB % (6 - b)))),
            );
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test("sequential typing batches use reciprocal coordinates", () => {
    converge(
      documentOf(),
      batch(insertion(2, "XY")),
      batch(insertion(1, "Q"), insertion(3, "R"), insertion(4, "S")),
    );
  });

  test(
    "a reciprocal join length includes earlier local insertions and deletions",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 3 }),
          fc.integer({ min: 1, max: 3 }),
          fc.constantFrom("X", "😀", "XYZ"),
          (firstPoint, secondPoint, text) => {
            const join = {
              type: "joinBlocks",
              story: "main",
              blockId: "00000001",
              nextBlockId: "00000002",
            } as const;
            const foreign = {
              ...sequenced(join),
              effects: [{ type: "joinBlocks", firstLength: 3 }],
            } as const satisfies SequencedBatch;
            const inSecond = {
              ...insertion(secondPoint, "Z"),
              at: at(secondPoint, "00000002"),
            };
            const source: Document = {
              package: {
                document: {
                  content: [
                    paragraph("00000001", "abc"),
                    paragraph("00000002", "def"),
                  ],
                },
              },
            };
            for (const first of [insertion(firstPoint, text), deletion(0, 1)]) {
              const delta = first.type === "insertText" ? text.length : -1;
              const ops = transformed(batch(first, inSecond), foreign);
              expect(ops.at(-1)).toEqual({
                ...inSecond,
                at: at(3 + delta + secondPoint, "00000002"),
              });
              expect(
                equalForStaleness(
                  apply(apply(source, [join]), ops),
                  apply(apply(source, [first, inSecond]), [join]),
                ),
              ).toBe(true);
            }
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "text and explicit splits converge including equal offsets",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 6 }),
          fc.integer({ min: 0, max: 6 }),
          (point, split) => {
            converge(
              documentOf(),
              batch({
                type: "splitBlock",
                at: at(split),
                newBlockId: "00000002",
                newHalf: "first",
              }),
              batch(insertion(point, "X")),
            );
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "direct ranges crossing a split map into both paragraphs",
    () => {
      fc.assert(
        fc.property(fc.integer({ min: 1, max: 5 }), (split) => {
          converge(
            documentOf(),
            batch({
              type: "splitBlock",
              at: at(split),
              newBlockId: "00000002",
              newHalf: "first",
            }),
            batch(deletion(0, 6)),
          );
        }),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "invalid cross-paragraph ranges cannot become valid through rebasing",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 6 }),
          fc.integer({ min: 0, max: 6 }),
          fc.integer({ min: 0, max: 6 }),
          fc.boolean(),
          (fromOffset, toOffset, splitOffset, reverse) => {
            const from = at(fromOffset, reverse ? "00000002" : "00000001");
            const to = at(toOffset, reverse ? "00000001" : "00000002");
            const invalidRanges = [
              { type: "deleteRange", from, to },
              { type: "setRunProps", from, to, patch: { bold: true } },
            ] as const satisfies readonly DocumentOp[];
            const split = {
              type: "splitBlock",
              at: at(splitOffset),
              newBlockId: "00000003",
              newHalf: "second",
            } as const satisfies DocumentOp;
            const document = {
              package: {
                document: {
                  content: [
                    paragraph("00000001", "abcdef"),
                    paragraph("00000002", "uvwxyz"),
                  ],
                },
              },
            } satisfies Document;
            for (const invalid of invalidRanges) {
              expect(applyDocumentOps(document, [invalid]).isErr()).toBe(true);
              for (const order of ["before", "after"] as const) {
                for (const tail of [
                  [],
                  [sequenced(split)],
                  [sequenced(insertion(0, "X"))],
                ]) {
                  const result = transformBatch(batch(invalid), tail, {
                    order,
                  });
                  expect(result.isErr()).toBe(true);
                  if (result.isErr())
                    expect(result.error.reason).toBe(
                      BATCH_REJECTION_REASONS.INVALID_OPERATION,
                    );
                }
                // Journal validation must run even with no incoming operations or an
                // earlier journal operation that would consume the incoming range.
                for (const incoming of [
                  batch(),
                  batch(deletion(0, 6)),
                  batch(insertion(0, "X")),
                ]) {
                  expect(
                    transformBatch(
                      incoming,
                      [sequenced(deletion(0, 6)), sequenced(invalid)],
                      {
                        order,
                      },
                    ).isErr(),
                  ).toBe(true);
                }
              }
              const sequencer = createSequencer(document);
              // The split's new identity lies in the actor's first lease block.
              sequencer.grantIdLease("actor").unwrap();
              expect(sequencer.submit(batch(split)).type).toBe("ack");
              const before = sequencer.document;
              const submission = { ...batch(invalid), opId: "invalid-range" };
              const rejected = sequencer.submit(submission);
              expect(rejected.type).toBe("reject");
              expect(sequencer.submit(submission)).toEqual(rejected);
              expect(equalForStaleness(sequencer.document, before)).toBe(true);
              expect(sequencer.headRev).toBe(1);
              expect(sequencer.broadcasts).toHaveLength(1);
            }
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "two explicit splits converge including the same boundary",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 0, max: 6 }),
          fc.integer({ min: 0, max: 6 }),
          (a, b) => {
            converge(
              documentOf(),
              batch({
                type: "splitBlock",
                at: at(a),
                newBlockId: "00000002",
                newHalf: "first",
              }),
              batch({
                type: "splitBlock",
                at: at(b),
                newBlockId: "00000003",
                newHalf: "first",
              }),
            );
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "run properties and inherited text converge at every range endpoint",
    () => {
      fc.assert(
        fc.property(fc.integer({ min: 0, max: 6 }), (point) => {
          converge(
            documentOf(),
            batch({
              type: "setRunProps",
              from: at(1),
              to: at(5),
              patch: { bold: true },
            }),
            batch(insertion(point, "X")),
          );
        }),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test("disjoint paragraph edits commute", () => {
    const op = {
      ...insertion(1, "X"),
      at: at(1, "00000002"),
    } satisfies DocumentOp;
    expect(transformed(batch(op), sequenced(insertion(3, "Y")))).toEqual([op]);
  });

  test("split halves and join survivors follow the real operation contract", () => {
    const split = {
      type: "splitBlock",
      at: at(3),
      newBlockId: "00000002",
      newHalf: "first",
    } as const;
    expect(transformed(batch(insertion(1, "X")), sequenced(split))).toEqual([
      { ...insertion(1, "X"), at: at(1, "00000002") },
    ]);
    expect(transformed(batch(insertion(5, "X")), sequenced(split))).toEqual([
      insertion(2, "X"),
    ]);
    const joined = {
      ...sequenced({
        type: "joinBlocks",
        story: "main",
        blockId: "00000001",
        nextBlockId: "00000002",
      }),
      effects: [{ type: "joinBlocks", firstLength: 3 }],
    } as const satisfies SequencedBatch;
    expect(transformed(batch(insertion(2, "X")), joined)).toEqual([
      { ...insertion(2, "X"), at: at(2, "00000002") },
    ]);
    const inSecond = { ...insertion(2, "X"), at: at(2, "00000002") };
    expect(transformed(batch(inSecond), joined)).toEqual([
      { ...inSecond, at: at(5, "00000002") },
    ]);
  });

  test("tracked deletes keep offsets; structural inverses keep preconditions", () => {
    const tracked = {
      ...deletion(1, 4),
      revision: { id: 10, author: "Actor", date: "2026-01-01T00:00:00Z" },
    } satisfies DocumentOp;
    expect(transformed(batch(insertion(5, "X")), sequenced(tracked))).toEqual([
      insertion(5, "X"),
    ]);
    expect(
      transformBatch(batch(insertion(2, "X")), [sequenced(tracked)]).isErr(),
    ).toBe(true);
    expect(
      transformBatch(batch(tracked), [sequenced(insertion(2, "X"))]).isErr(),
    ).toBe(true);
    const inverse = {
      type: "deleteRange",
      from: at(1),
      to: at(2),
      expected: {
        content: [{ type: "run", content: [{ type: "text", text: "b" }] }],
        openStart: 0,
        openEnd: 0,
      },
    } satisfies DocumentOp;
    expect(transformed(batch(inverse), sequenced(insertion(3, "X")))).toEqual([
      inverse,
    ]);
    expect(
      transformBatch(batch(inverse), [sequenced(insertion(1, "X"))]).isErr(),
    ).toBe(true);
  });

  test("a structural inverse passes a split away from its range and a join of its paragraph", () => {
    // The undo of typing "c": a deletion that expects the typed text.
    const typed = applyBatch(
      {
        package: {
          document: {
            content: [
              paragraph("00000001", "xyz"),
              paragraph("00000002", "abdef"),
            ],
          },
        },
      },
      [{ ...insertion(2, "c"), at: at(2, "00000002") }],
    );
    if (typed.isErr()) throw typed.error;
    const { document } = typed.value;
    const [inverse] = typed.value.inverse;
    if (inverse?.type !== "deleteRange" || inverse.expected === undefined)
      throw new Error("Typing must record a structural inverse.");
    const split = (offset: number) =>
      ({
        type: "splitBlock",
        at: at(offset, "00000002"),
        newBlockId: "00000009",
        newHalf: "second",
      }) as const satisfies DocumentOp;
    const join = {
      type: "joinBlocks",
      story: "main",
      blockId: "00000001",
      nextBlockId: "00000002",
      survivor: "first",
    } as const satisfies DocumentOp;
    // Both orders give one document: the range's content is the same.
    for (const over of [split(5), split(0), join]) {
      const remote = applyBatch(document, [over]);
      if (remote.isErr()) throw remote.error;
      const moved = transformBatch(batch(inverse), [
        { ...sequenced(over), effects: remote.value.effects },
      ]);
      if (moved.isErr()) throw moved.error;
      const left = apply(remote.value.document, moved.value.ops);
      const local = applyBatch(document, [inverse]);
      if (local.isErr()) throw local.error;
      const past = transformBatch(
        batch(over),
        [{ ...sequenced(inverse), effects: local.value.effects }],
        { order: "before" },
      );
      if (past.isErr()) throw past.error;
      const right = apply(local.value.document, past.value.ops);
      expect(equalForStaleness(left, right)).toBe(true);
    }
    // A split at or inside the range still changes what it expects.
    for (const offset of [2, 3]) {
      const remote = applyBatch(document, [split(offset)]);
      if (remote.isErr()) throw remote.error;
      expect(
        transformBatch(batch(inverse), [
          { ...sequenced(split(offset)), effects: remote.value.effects },
        ]).isErr(),
      ).toBe(true);
    }
  });
});

describe("insertions at a gap where a record was cut", () => {
  const typed = (
    text: string,
    revision?: { id: number; author: string; date: string },
  ) =>
    ({
      type: "insertContent",
      at: at(2),
      slice: {
        content: [{ type: "run", content: [{ type: "text", text }] }],
        openStart: 0,
        openEnd: 0,
      },
      ...(revision === undefined ? {} : { revision }),
    }) as const satisfies DocumentOp;
  const stamp = { id: 30, author: "Actor", date: "2026-01-01T00:00:00Z" };
  const cutting = (op: DocumentOp): SequencedBatch => ({
    ...sequenced(op),
    effects: [{ type: "insertContent", cutIdentified: true }],
  });

  test("two direct insertions refuse: the cut record's second half takes either id", () => {
    expect(
      transformBatch(batch(typed("x")), [cutting(typed("y"))]).isErr(),
    ).toBe(true);
  });

  test("a tracked insertion passes, after the one sequenced first", () => {
    for (const [op, over] of [
      [typed("x", stamp), typed("y")],
      [typed("x"), typed("y", stamp)],
    ] as const) {
      const moved = transformBatch(batch(op), [cutting(over)]);
      if (moved.isErr()) throw moved.error;
      expect(moved.value.ops).toEqual([{ ...op, at: at(3) }]);
    }
  });
});

describe("concurrent edits meeting an identified record", () => {
  // "Pa" is a tracked deletion: inserting between "P" and "a" cuts it, its
  // second half taking the insertion's new id.
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "deletion",
                info: { id: 5, author: "Actor", date: "2026-01-01T00:00:00Z" },
                content: [
                  { type: "run", content: [{ type: "text", text: "Pa" }] },
                ],
              },
              { type: "run", content: [{ type: "text", text: "rest" }] },
            ],
          },
        ],
      },
    },
  };
  const typed = {
    type: "insertContent",
    at: { ...at(1), zeroWidthBefore: 0 },
    slice: {
      content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
      openStart: 0,
      openEnd: 0,
    },
    newIds: { revision: [12], control: [] },
  } as const satisfies DocumentOp;
  const inOrder = (first: DocumentOp, second: DocumentOp) => {
    const applied = applyBatch(document, [first]);
    if (applied.isErr()) throw applied.error;
    return transformBatch(batch(second), [
      { ...sequenced(first), effects: applied.value.effects },
    ]);
  };

  test("a deletion reaching back from the cut refuses in either order", () => {
    // Deleting "P" first leaves "a" deletion 5; cutting first leaves it 12.
    for (const removed of [deletion(0, 1), deletion(0, 2)]) {
      const message =
        "Deleting up to an insertion that cuts an identified record needs explicit affinity.";
      const refused = [inOrder(typed, removed), inOrder(removed, typed)];
      for (const result of refused) {
        expect(result.isErr()).toBe(true);
        if (result.isErr()) expect(result.error.message).toBe(message);
      }
    }
  });

  test("two insertions of one identified record refuse in either order", () => {
    const record = (offset: number) =>
      ({
        type: "insertContent",
        at: { ...at(offset), zeroWidthBefore: 0 },
        slice: {
          content: [
            {
              type: "insertion",
              info: { id: 7, author: "Actor", date: "2026-01-01T00:00:00Z" },
              content: [
                { type: "run", content: [{ type: "text", text: "y" }] },
              ],
            },
          ],
          openStart: 0,
          openEnd: 0,
        },
      }) as const satisfies DocumentOp;
    for (const [first, second] of [
      [record(2), record(2)],
      [record(2), record(4)],
      [record(4), record(2)],
    ] as const) {
      const result = inOrder(first, second);
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.message).toBe(
          "Concurrent insertions of one identified record conflict.",
        );
    }
  });

  test("a deletion starting at the cut passes in either order", () => {
    const removed = deletion(1, 2);
    expect(inOrder(typed, removed).isOk()).toBe(true);
    expect(inOrder(removed, typed).isOk()).toBe(true);
  });
});

describe("property and structural policies", () => {
  test("per-key paragraph and run property decisions converge", () => {
    for (const type of ["setParagraphProps", "setRunProps"] as const) {
      const first: DocumentOp =
        type === "setParagraphProps"
          ? {
              type,
              story: "main",
              blockId: "00000001",
              patch: { spaceAfter: 10 },
            }
          : { type, from: at(0), to: at(6), patch: { bold: true } };
      const second: DocumentOp =
        type === "setParagraphProps"
          ? {
              type,
              story: "main",
              blockId: "00000001",
              patch: { spaceAfter: 20 },
            }
          : {
              type,
              from: at(0),
              to: at(6),
              patch: { bold: false, italic: true },
            };
      converge(documentOf(), batch(first), batch(second));
    }
  });

  test("partially overlapping competing run keys require captured boundaries", () => {
    const left = {
      type: "setRunProps",
      from: at(1),
      to: at(4),
      patch: { bold: true },
    } as const;
    const right = {
      type: "setRunProps",
      from: at(2),
      to: at(5),
      patch: { bold: false },
    } as const;
    expect(transformBatch(batch(left), [sequenced(right)]).isErr()).toBe(true);
    expect(transformBatch(batch(right), [sequenced(left)]).isErr()).toBe(true);
  });

  test("block insertions maintain sequencer order on both anchor sides", () => {
    for (const type of ["before", "after"] as const) {
      const first = batch({
        type: "insertBlocks",
        story: "main",
        at: { type, blockId: "00000001" },
        blocks: [paragraph("00000002", "first")],
      });
      const second = batch({
        type: "insertBlocks",
        story: "main",
        at: { type, blockId: "00000001" },
        blocks: [paragraph("00000003", "second")],
      });
      converge(documentOf(), first, second);
    }
  });

  test("revision decisions are idempotent and opposing decisions reject", () => {
    const resolve = {
      type: "resolveRevision",
      story: "main",
      revisionIds: [10],
      decision: "accept",
    } as const;
    expect(transformed(batch(resolve), sequenced(resolve))).toEqual([resolve]);
    expect(
      transformBatch(batch({ ...resolve, decision: "reject" }), [
        sequenced(resolve),
      ]).isErr(),
    ).toBe(true);
  });

  test("a decision passes a direct join that states both halves", () => {
    const resolve = {
      type: "resolveRevision",
      story: "main",
      revisionIds: [10],
      decision: "accept",
    } as const;
    const join = {
      type: "joinBlocks",
      story: "main",
      blockId: "00000002",
      nextBlockId: "00000001",
      depth: 2,
      survivor: "second",
      expectedRetired: {},
      expectedSurvivor: {},
    } as const;
    expect(transformed(batch(resolve), sequenced(join))).toEqual([resolve]);
    // Unstated halves may hold a property change the join moves.
    const { expectedRetired: _retired, ...unstated } = join;
    expect(transformBatch(batch(resolve), [sequenced(unstated)]).isErr()).toBe(
      true,
    );
    const revision = { id: 20, author: "Actor", date: "2026-01-01T00:00:00Z" };
    expect(
      transformBatch(batch(resolve), [
        sequenced({ ...join, depth: 0, revision }),
      ]).isErr(),
    ).toBe(true);
  });

  test("a decision elsewhere moves past each operation of a tracked batch", () => {
    const resolve = {
      type: "resolveRevision",
      story: "main",
      revisionIds: [10],
      decision: "accept",
    } as const;
    const revision = { id: 20, author: "Actor", date: "2026-01-01T00:00:00Z" };
    const replacement = batch(
      { ...deletion(0, 1), revision },
      { ...insertion(1, "q"), revision },
    );
    const decided = (blockIds: string[]) =>
      transformBatch(replacement, [
        {
          ...sequenced(resolve),
          effects: [{ type: "touchedBlocks", blockIds }],
        },
      ]);
    // Moving the decision past the deletion to meet the insertion: the
    // paragraphs it decided hold none of the batch's edits.
    const passed = decided(["00000002"]);
    if (passed.isErr()) throw passed.error;
    expect(passed.value.ops).toEqual(replacement.ops);
    // In a paragraph the decision changed, the edit is refused.
    expect(decided(["00000001"]).isErr()).toBe(true);
  });

  test("numbering, deleted targets and structural overlaps reject", () => {
    const props = {
      type: "setParagraphProps",
      story: "main",
      blockId: "00000001",
      patch: { numPr: null },
    } as const;
    expect(transformBatch(batch(props), [sequenced(props)]).isErr()).toBe(true);
    expect(
      transformBatch(batch(insertion(1, "X")), [
        sequenced({
          type: "deleteBlocks",
          story: "main",
          blockIds: ["00000001"],
        }),
      ]).isErr(),
    ).toBe(true);
    expect(
      transformBatch(
        batch({
          type: "joinBlocks",
          story: "main",
          blockId: "00000001",
          nextBlockId: "00000002",
        }),
        [
          sequenced({
            type: "splitBlock",
            at: at(2),
            newBlockId: "00000003",
            newHalf: "second",
          }),
        ],
      ).isErr(),
    ).toBe(true);
  });

  test("captured structural footprints commute only with unaffected targets", () => {
    const removed = {
      ...sequenced({
        type: "deleteBlocks",
        story: "main",
        blockIds: ["00000002"],
      }),
      effects: [{ type: "touchedBlocks", blockIds: ["00000002", "00000003"] }],
    } as const satisfies SequencedBatch;
    expect(transformed(batch(insertion(1, "X")), removed)).toEqual([
      insertion(1, "X"),
    ]);
    const affected = { ...insertion(1, "X"), at: at(1, "00000003") };
    expect(transformBatch(batch(affected), [removed]).isErr()).toBe(true);
  });
});

test("story lifecycle operations pass text edits of other paragraphs and conflict with each other", () => {
  const lifecycle = new Set([
    DOCUMENT_OP_TYPES.ADD_NOTE,
    DOCUMENT_OP_TYPES.REMOVE_NOTE,
    DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
    DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
    DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
    DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
  ]);
  const exercised = new Set<string>();
  const elsewhere = { ...insertion(0, "X"), at: at(0, "0000ABCD") };
  for (const { op } of envelopes()) {
    if (![...lifecycle].some((type) => type === op.type)) continue;
    exercised.add(op.type);
    for (const result of [
      transformBatch(batch(op), [sequenced(elsewhere)]),
      transformBatch(batch(elsewhere), [sequenced(op)]),
    ])
      expect(result.isOk()).toBe(true);
    const same = transformBatch(batch(op), [sequenced(op)]);
    expect(same.isErr()).toBe(true);
    if (same.isErr())
      expect([
        BATCH_REJECTION_REASONS.PART_CONFLICT,
        BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
      ]).toContain(same.error.reason);
  }
  expect([...exercised].toSorted()).toEqual([...lifecycle].toSorted());
});
