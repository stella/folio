import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { equalForStaleness } from "../equality";
import { DOCUMENT_OP_SCHEMA_VERSION, type DocumentOp } from "../types";
import { BATCH_REJECTION_REASONS, type DocumentBatch, type SequencedBatch } from "./envelope";
import { createSequencer } from "./sequencer";
import { transformBatch } from "./transform";

const at = (offset: number, blockId = "00000001") => ({ story: "main", blockId, offset }) as const;
const insertion = (offset: number, text: string) =>
  ({ type: "insertText", at: at(offset), text, runProps: "inherit" }) as const satisfies DocumentOp;
const deletion = (from: number, to: number) =>
  ({ type: "deleteRange", from: at(from), to: at(to) }) as const satisfies DocumentOp;
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
const converge = (document: Document, first: DocumentBatch, second: DocumentBatch) => {
  const firstSequenced = { ...first, revision: 1 };
  const secondSequenced = { ...second, revision: 1 };
  const left = apply(apply(document, first.ops), transformed(second, firstSequenced));
  const right = apply(apply(document, second.ops), transformed(first, secondSequenced, "before"));
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
            converge(documentOf(), batch(insertion(left, text)), batch(insertion(right, "Z")));
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
            converge(documentOf(), batch(deletion(from, to)), batch(insertion(point, "X")));
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
            const inSecond = { ...insertion(secondPoint, "Z"), at: at(secondPoint, "00000002") };
            const source: Document = {
              package: {
                document: { content: [paragraph("00000001", "abc"), paragraph("00000002", "def")] },
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
            batch({ type: "splitBlock", at: at(split), newBlockId: "00000002", newHalf: "first" }),
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
                  content: [paragraph("00000001", "abcdef"), paragraph("00000002", "uvwxyz")],
                },
              },
            } satisfies Document;
            for (const invalid of invalidRanges) {
              expect(applyDocumentOps(document, [invalid]).isErr()).toBe(true);
              for (const order of ["before", "after"] as const) {
                for (const tail of [[], [sequenced(split)], [sequenced(insertion(0, "X"))]]) {
                  const result = transformBatch(batch(invalid), tail, { order });
                  expect(result.isErr()).toBe(true);
                  if (result.isErr())
                    expect(result.error.reason).toBe(BATCH_REJECTION_REASONS.INVALID_OPERATION);
                }
                // Journal validation must run even with no incoming operations or an
                // earlier journal operation that would consume the incoming range.
                for (const incoming of [batch(), batch(deletion(0, 6)), batch(insertion(0, "X"))]) {
                  expect(
                    transformBatch(incoming, [sequenced(deletion(0, 6)), sequenced(invalid)], {
                      order,
                    }).isErr(),
                  ).toBe(true);
                }
              }
              const sequencer = createSequencer(document);
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
        fc.property(fc.integer({ min: 0, max: 6 }), fc.integer({ min: 0, max: 6 }), (a, b) => {
          converge(
            documentOf(),
            batch({ type: "splitBlock", at: at(a), newBlockId: "00000002", newHalf: "first" }),
            batch({ type: "splitBlock", at: at(b), newBlockId: "00000003", newHalf: "first" }),
          );
        }),
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
            batch({ type: "setRunProps", from: at(1), to: at(5), patch: { bold: true } }),
            batch(insertion(point, "X")),
          );
        }),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(10_000),
  );

  test("disjoint paragraph edits commute", () => {
    const op = { ...insertion(1, "X"), at: at(1, "00000002") } satisfies DocumentOp;
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
    expect(transformed(batch(insertion(5, "X")), sequenced(split))).toEqual([insertion(2, "X")]);
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
    expect(transformed(batch(inSecond), joined)).toEqual([{ ...inSecond, at: at(5, "00000002") }]);
  });

  test("tracked deletes keep offsets; structural inverses keep preconditions", () => {
    const tracked = {
      ...deletion(1, 4),
      revision: { id: 10, author: "Actor", date: "2026-01-01T00:00:00Z" },
    } satisfies DocumentOp;
    expect(transformed(batch(insertion(5, "X")), sequenced(tracked))).toEqual([insertion(5, "X")]);
    expect(transformBatch(batch(insertion(2, "X")), [sequenced(tracked)]).isErr()).toBe(true);
    expect(transformBatch(batch(tracked), [sequenced(insertion(2, "X"))]).isErr()).toBe(true);
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
    expect(transformed(batch(inverse), sequenced(insertion(3, "X")))).toEqual([inverse]);
    expect(transformBatch(batch(inverse), [sequenced(insertion(1, "X"))]).isErr()).toBe(true);
  });
});

describe("property and structural policies", () => {
  test("per-key paragraph and run property decisions converge", () => {
    for (const type of ["setParagraphProps", "setRunProps"] as const) {
      const first: DocumentOp =
        type === "setParagraphProps"
          ? { type, story: "main", blockId: "00000001", patch: { spaceAfter: 10 } }
          : { type, from: at(0), to: at(6), patch: { bold: true } };
      const second: DocumentOp =
        type === "setParagraphProps"
          ? { type, story: "main", blockId: "00000001", patch: { spaceAfter: 20 } }
          : { type, from: at(0), to: at(6), patch: { bold: false, italic: true } };
      converge(documentOf(), batch(first), batch(second));
    }
  });

  test("partially overlapping competing run keys require captured boundaries", () => {
    const left = { type: "setRunProps", from: at(1), to: at(4), patch: { bold: true } } as const;
    const right = { type: "setRunProps", from: at(2), to: at(5), patch: { bold: false } } as const;
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
      transformBatch(batch({ ...resolve, decision: "reject" }), [sequenced(resolve)]).isErr(),
    ).toBe(true);
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
        sequenced({ type: "deleteBlocks", story: "main", blockIds: ["00000001"] }),
      ]).isErr(),
    ).toBe(true);
    expect(
      transformBatch(
        batch({ type: "joinBlocks", story: "main", blockId: "00000001", nextBlockId: "00000002" }),
        [sequenced({ type: "splitBlock", at: at(2), newBlockId: "00000003", newHalf: "second" })],
      ).isErr(),
    ).toBe(true);
  });

  test("captured structural footprints commute only with unaffected targets", () => {
    const removed = {
      ...sequenced({ type: "deleteBlocks", story: "main", blockIds: ["00000002"] }),
      effects: [{ type: "touchedBlocks", blockIds: ["00000002", "00000003"] }],
    } as const satisfies SequencedBatch;
    expect(transformed(batch(insertion(1, "X")), removed)).toEqual([insertion(1, "X")]);
    const affected = { ...insertion(1, "X"), at: at(1, "00000003") };
    expect(transformBatch(batch(affected), [removed]).isErr()).toBe(true);
  });
});
