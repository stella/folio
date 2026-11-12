// Coverage stated without the draw. scripts/lib/property-draw-coverage.ts
// reports nothing here.
import { expect, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample, coverageSample } from "../property-coverage";
import { assertProperty } from "../property-testing";

const KINDS = ["a", "b", "c"] as const;
type Kind = (typeof KINDS)[number];
type Tally = Map<Kind, number>;

const count = (tally: Tally, kind: Kind): void => {
  tally.set(kind, (tally.get(kind) ?? 0) + 1);
};

const expectEveryKind = (tally: Tally): void => {
  for (const kind of KINDS) expect(tally.get(kind) ?? 0).toBeGreaterThan(0);
};

const law =
  (tally: Tally) =>
  (kind: Kind): void => {
    count(tally, kind);
  };

test("the property searches with a tally nobody reads", () => {
  assertProperty(fc.property(fc.constantFrom(...KINDS), law(new Map())), {
    id: "law",
  });
});

test("the same check reaches every kind on a fixed sample", () => {
  const tally: Tally = new Map();
  const check = law(tally);
  for (const kind of coverageSample(fc.constantFrom(...KINDS), 30)) check(kind);
  expectEveryKind(tally);
});

test("fixed examples per kind, checked where they are written", () => {
  const examples = KINDS.map((kind): [Kind] => [kind]);
  expect(examples.map(([kind]) => kind)).toEqual([...KINDS]);
  assertProperty(
    fc.property(fc.constantFrom(...KINDS), (kind) => KINDS.includes(kind)),
    { id: "examples", examples },
  );
});

test("a run under a seed of its own replays the same cases everywhere", () => {
  const seen = new Set<Kind>();
  assertProperty(
    fc.property(fc.constantFrom(...KINDS), (kind) => {
      seen.add(kind);
    }),
    { id: "seeded", seed: 7, numRuns: 30 },
  );
  expect(seen.size).toBe(KINDS.length);
});

test("a sample that names its seed", () => {
  const seen = new Set<number>();
  for (const value of fc.sample(fc.nat(), { seed: 3, numRuns: 50 }))
    seen.add(value % 3);
  expect(seen.size).toBe(3);
});

test("nothing refused: a truncated run cannot fail what stays empty", () => {
  const refusals: string[] = [];
  assertProperty(
    fc.property(fc.nat(), (value) => {
      if (value < 0) refusals.push(String(value));
    }),
    { id: "refusals" },
  );
  expect(refusals).toEqual([]);
  expect(refusals).toHaveLength(0);
});

test("a name the predicate shadows is a different binding", () => {
  const applied = [1, 2, 3];
  assertProperty(
    fc.property(fc.nat(), (value) => {
      let applied = 0;
      applied += value;
      return applied >= 0;
    }),
    { id: "shadowed" },
  );
  expect(applied.length).toBeGreaterThan(0);
});

const lawOver = (seen: Set<Kind>) =>
  fc.property(fc.constantFrom(...KINDS), (kind) => {
    seen.add(kind);
  });

test("the property searches; its own predicate reaches every kind on a fixed sample", () => {
  assertProperty(lawOver(new Set()), { id: "built" });
  const seen = new Set<Kind>();
  checkCoverageSample(lawOver(seen), 30);
  expect(seen.size).toBe(KINDS.length);
});
