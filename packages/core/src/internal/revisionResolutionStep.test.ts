import { describe, expect, test } from "bun:test";
import { Fragment, Schema, Slice } from "prosemirror-model";
import {
  AttrStep,
  Mapping,
  RemoveMarkStep,
  ReplaceStep,
  Step,
  StepMap,
} from "prosemirror-transform";

import { composeRevisionResolutionMaps, RevisionResolutionStep } from "./revisionResolutionStep";

const schema = new Schema({
  nodes: {
    doc: { content: "paragraph+" },
    paragraph: { attrs: { id: { default: null } }, content: "text*", group: "block" },
    text: { group: "inline" },
  },
  marks: { insertion: {} },
});

const paragraph = (id: string, text: string) =>
  schema.node("paragraph", { id }, [schema.text(text)]);

const source = schema.node("doc", null, [
  schema.node("paragraph", { id: "one" }, [schema.text("abc", [schema.mark("insertion")])]),
  paragraph("two", "def"),
  paragraph("three", "ghi"),
]);

const edits = [
  new RemoveMarkStep(1, 4, schema.mark("insertion")),
  new ReplaceStep(7, 8, Slice.empty),
  new AttrStep(9, "id", "resolved-three"),
];

const replay = (doc: typeof source, steps: readonly Step[]) => {
  let current = doc;
  for (const step of steps) {
    const result = step.apply(current);
    if (!result.doc) {
      throw new Error(result.failed ?? "Step failed");
    }
    current = result.doc;
  }
  return current;
};

describe("RevisionResolutionStep", () => {
  test.each([
    [-1, 1, 0],
    [0, -1, 0],
    [0, 1, -1],
    [3, 2, 0, 4, 1, 0],
    [3, 1, 0, 2, 1, 0],
    [Number.MAX_SAFE_INTEGER, 1, 0],
    [0, 0, Number.MAX_SAFE_INTEGER, 0, 0, 1],
  ])("rejects malformed serialized position ranges %j", (mapRanges) => {
    expect(() =>
      Step.fromJSON(schema, { stepType: "folioRevisionResolution", steps: [], mapRanges }),
    ).toThrow(RangeError);
  });

  test("composes inline and structural ranges in source coordinates", () => {
    const inline = new StepMap([7, 1, 0, 15, 2, 0]);
    const structural = new StepMap([7, 2, 0, 12, 2, 1]);
    const composed = composeRevisionResolutionMaps(inline, structural);
    const sequential = new Mapping();
    sequential.appendMap(inline);
    sequential.appendMap(structural);
    for (const position of [0, 1, 5, 10, 11, 12, 18, 20]) {
      expect(composed.map(position)).toBe(sequential.map(position));
    }
    const ranges: Array<[number, number, number, number]> = [];
    composed.forEach((oldStart, oldEnd, newStart, newEnd) => {
      ranges.push([oldStart, oldEnd, newStart, newEnd]);
    });
    expect(ranges).toEqual([
      [7, 10, 7, 7],
      [13, 17, 10, 11],
    ]);
  });

  test("composes a boundary join after deleting its neighboring block", () => {
    const composed = composeRevisionResolutionMaps(new StepMap([8, 28, 0]), new StepMap([7, 2, 0]));
    const ranges: number[] = [];
    composed.forEach((oldStart, oldEnd, _newStart, newEnd) => {
      ranges.push(oldStart, oldEnd - oldStart, newEnd);
    });
    expect(ranges).toEqual([7, 30, 7]);
  });

  test("composed maps keep all surviving positions over varied deletion layouts", () => {
    const mapFor = (size: number, seed: number) => {
      const ranges: number[] = [];
      let removed = 0;
      for (let position = 1; position < size - 1; position += 1) {
        if ((position * 17 + seed * 13) % 11 !== 0) {
          continue;
        }
        const length = Math.min(1 + ((position + seed) % 3), size - position - 1);
        ranges.push(position, length, 0);
        removed += length;
        position += length - 1;
      }
      return { map: new StepMap(ranges), size: size - removed };
    };
    for (let seed = 0; seed < 30; seed += 1) {
      const first = mapFor(40, seed);
      const second = mapFor(first.size, seed + 31);
      const composed = composeRevisionResolutionMaps(first.map, second.map);
      const sequential = new Mapping();
      sequential.appendMap(first.map);
      sequential.appendMap(second.map);
      for (let position = 0; position <= 40; position += 1) {
        for (const assoc of [-1, 1]) {
          if (!sequential.mapResult(position, assoc).deleted) {
            expect(composed.map(position, assoc)).toBe(sequential.map(position, assoc));
          }
        }
      }
    }
  });

  test("applies, serializes, and inverts granular changes", () => {
    const resolved = replay(source, edits);
    const step = new RevisionResolutionStep({
      beforeDoc: source,
      afterDoc: resolved,
      steps: edits,
      positionMap: new StepMap([7, 1, 0]),
    });
    expect(step.apply(source).doc).toBe(resolved);
    expect(Step.fromJSON(schema, step.toJSON()).apply(source).doc?.eq(resolved)).toBe(true);
    expect(step.getMap().map(12)).toBe(11);
    expect(step.getMap().map(2)).toBe(2);

    const inverse = step.invert(source);
    expect(inverse.apply(resolved).doc).toBe(source);
    expect(Step.fromJSON(schema, inverse.toJSON()).apply(resolved).doc?.eq(source)).toBe(true);
    expect(inverse.getMap().map(11)).toBe(12);
  });

  test("maps independent edits around an earlier insertion", () => {
    const resolved = replay(source, edits);
    const step = new RevisionResolutionStep({
      beforeDoc: source,
      afterDoc: resolved,
      steps: edits,
      positionMap: new StepMap([7, 1, 0]),
    });
    const prefix = paragraph("prefix", "z");
    const insertion = new ReplaceStep(0, 0, new Slice(Fragment.from(prefix), 0, 0));
    const external = insertion.apply(source).doc;
    if (!external) {
      throw new Error("Prefix insertion failed");
    }
    const mapping = new Mapping();
    mapping.appendMap(insertion.getMap());
    const rebased = step.map(mapping);
    expect(rebased).not.toBeNull();
    expect(
      rebased
        ?.apply(external)
        .doc?.eq(schema.node("doc", null, [prefix, ...resolved.content.content])),
    ).toBe(true);
  });

  test("preserves a concurrent insertion between resolved changes", () => {
    const resolved = replay(source, edits);
    const step = new RevisionResolutionStep({
      beforeDoc: source,
      afterDoc: resolved,
      steps: edits,
      positionMap: new StepMap([7, 1, 0]),
    });
    const text = new Slice(Fragment.from(schema.text("X")), 0, 0);
    const insertion = new ReplaceStep(9, 9, text);
    const external = insertion.apply(source).doc;
    if (!external) {
      throw new Error("Concurrent insertion failed");
    }
    const mapping = new Mapping();
    mapping.appendMap(insertion.getMap());
    const rebased = step.map(mapping);
    const expected = new ReplaceStep(8, 8, text).apply(resolved).doc;
    expect(rebased?.apply(external).doc?.eq(expected ?? source)).toBe(true);
    expect(rebased?.getMap().map(13)).toBe(12);
  });
  test("rebases overlapping removal without dropping independent mark resolution", () => {
    const deletion = new ReplaceStep(6, 9, Slice.empty);
    const original = [edits[0]!, deletion];
    const resolved = replay(source, original);
    const step = new RevisionResolutionStep({
      beforeDoc: source,
      afterDoc: resolved,
      steps: original,
      positionMap: deletion.getMap(),
    });
    const externalEdit = new ReplaceStep(7, 8, Slice.empty);
    const external = externalEdit.apply(source).doc;
    if (!external) throw new Error("Concurrent deletion failed");
    const rebased = step.map(externalEdit.getMap());
    expect(rebased).not.toBeNull();
    expect(rebased?.apply(external).doc?.eq(resolved)).toBe(true);
    expect(rebased?.getMap().map(external.content.size, 1)).toBe(resolved.content.size);
  });
});
