import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { ParagraphBlock, Run } from "../types";
import { withFakeTextMeasure } from "./__tests__/fakeTextMeasure";
import { clearAllCaches } from "./cache";
import { measureBlock } from "./measureBlocks";
import { noteParagraphOrigin } from "./paragraphKeyMemo";

type RunSpec =
  | {
      kind: "text";
      text: string;
      bold: boolean;
      italic: boolean;
      fontSize: 9 | 11 | 16;
    }
  | { kind: "tab" }
  | { kind: "lineBreak" };

type Edit =
  | { kind: "insert"; run: number; at: number; char: string }
  | { kind: "delete"; run: number; at: number }
  | { kind: "bold"; run: number }
  | { kind: "size"; run: number; fontSize: 9 | 11 | 16 }
  | { kind: "move"; by: number };

const textRun = fc.record({
  kind: fc.constant("text" as const),
  text: fc.string({
    unit: fc.constantFrom("a", "B", " ", "w", "-"),
    maxLength: 40,
  }),
  bold: fc.boolean(),
  italic: fc.boolean(),
  fontSize: fc.constantFrom(9 as const, 11 as const, 16 as const),
});
const runSpec: fc.Arbitrary<RunSpec> = fc.oneof(
  { weight: 6, arbitrary: textRun },
  { weight: 1, arbitrary: fc.constant({ kind: "tab" as const }) },
  { weight: 1, arbitrary: fc.constant({ kind: "lineBreak" as const }) },
);
const edit: fc.Arbitrary<Edit> = fc.oneof(
  fc.record({
    kind: fc.constant("insert" as const),
    run: fc.nat(),
    at: fc.nat(),
    char: fc.constantFrom("a", "B", " "),
  }),
  fc.record({
    kind: fc.constant("delete" as const),
    run: fc.nat(),
    at: fc.nat(),
  }),
  fc.record({ kind: fc.constant("bold" as const), run: fc.nat() }),
  fc.record({
    kind: fc.constant("size" as const),
    run: fc.nat(),
    fontSize: fc.constantFrom(9 as const, 11 as const, 16 as const),
  }),
  fc.record({
    kind: fc.constant("move" as const),
    by: fc.integer({ min: -5, max: 40 }),
  }),
);

const applyEdit = (specs: RunSpec[], change: Edit): RunSpec[] => {
  if (change.kind === "move" || specs.length === 0) return specs;
  const index = change.run % specs.length;
  const spec = specs[index];
  if (spec?.kind !== "text") return specs;
  const offset = (at: number) => at % (spec.text.length + 1);
  const next = ((): RunSpec => {
    switch (change.kind) {
      case "insert": {
        const at = offset(change.at);
        return {
          ...spec,
          text: `${spec.text.slice(0, at)}${change.char}${spec.text.slice(at)}`,
        };
      }
      case "delete": {
        const at = offset(change.at);
        return {
          ...spec,
          text: `${spec.text.slice(0, at)}${spec.text.slice(at + 1)}`,
        };
      }
      case "bold":
        return { ...spec, bold: !spec.bold };
      case "size":
        return { ...spec, fontSize: change.fontSize };
    }
  })();
  return specs.map((value, position) => (position === index ? next : value));
};

/** The block of `specs` at `start`, its runs at their positions. */
const blockAt = (specs: RunSpec[], start: number): ParagraphBlock => {
  let position = start + 1;
  const runs = specs.map((spec): Run => {
    const size = spec.kind === "text" ? spec.text.length : 1;
    const pmStart = position;
    position += size;
    return { ...spec, pmStart, pmEnd: position };
  });
  return {
    kind: "paragraph",
    id: "p",
    pmStart: start,
    pmEnd: position + 1,
    runs,
  };
};

const measureOf = (block: ParagraphBlock, width: number) =>
  measureBlock(block, width);

test(
  "a cached paragraph measure equals a fresh one across edits and moves",
  async () => {
    await assertProperty(
      fc.property(
        fc.array(runSpec, { minLength: 1, maxLength: 6 }),
        fc.array(edit, { maxLength: 12 }),
        fc.constantFrom(60, 140, 400),
        (initial, edits, width) => {
          withFakeTextMeasure(() => {
            clearAllCaches();
            let specs = initial;
            let start = 0;
            let previous = blockAt(specs, start);
            measureOf(previous, width);
            for (const change of edits) {
              specs = applyEdit(specs, change);
              if (change.kind === "move")
                start = Math.max(0, start + change.by);
              const block = blockAt(specs, start);
              // The layout pass says which block a paragraph was moved from.
              noteParagraphOrigin(block, previous);
              const cached = measureOf(block, width);
              clearAllCaches();
              expect(cached).toEqual(measureOf(blockAt(specs, start), width));
              // Warm the caches again with what the edit started from.
              measureOf(previous, width);
              measureOf(block, width);
              previous = block;
            }
            clearAllCaches();
          });
        },
      ),
      { numRuns: 60, id: "cached paragraph measures equal fresh ones" },
    );
  },
  propertyTestTimeout(30_000),
);
