// Coverage asserted over what a property happened to draw. Every line marked
// `caught` must be reported by scripts/lib/property-draw-coverage.ts.
import { expect, test } from "bun:test";
import fc from "fast-check";

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

test("a tally a helper counts into, checked by a helper", () => {
  const tally: Tally = new Map();
  assertProperty(
    fc.property(fc.constantFrom(...KINDS), (kind) => {
      count(tally, kind);
    }),
    { id: "helper tally" },
  );
  expectEveryKind(tally); // caught
});

test("a counter", () => {
  let applied = 0;
  assertProperty(
    fc.property(fc.nat(), (value) => {
      if (value % 2 === 0) applied += 1;
    }),
    { id: "counter" },
  );
  expect(applied).toBeGreaterThan(10); // caught
});

test("a set of kinds seen", () => {
  const seen = new Set<Kind>();
  assertProperty(
    fc.property(fc.constantFrom(...KINDS), (kind) => {
      seen.add(kind);
    }),
    { id: "seen" },
  );
  expect([...seen].toSorted()).toEqual([...KINDS]); // caught
});

test("values read back from a tally", () => {
  const tally: Tally = new Map();
  assertProperty(
    fc.property(fc.constantFrom(...KINDS), (kind) => {
      count(tally, kind);
    }),
    { id: "read back" },
  );
  for (const applied of tally.values()) expect(applied).toBeGreaterThan(0); // caught
});

test("a callback among the arbitraries", () => {
  let mapped = 0;
  assertProperty(
    fc.property(
      fc.nat().map((value) => {
        mapped += 1;
        return value;
      }),
      () => true,
    ),
    { id: "mapped" },
  );
  expect(mapped).toBeGreaterThan(0); // caught
});

const law =
  (tally: Tally) =>
  (kind: Kind): void => {
    count(tally, kind);
  };

test("a predicate a factory builds over the tally", () => {
  const tally: Tally = new Map();
  assertProperty(fc.property(fc.constantFrom(...KINDS), law(tally)), {
    id: "factory",
  });
  expect(tally.size).toBe(KINDS.length); // caught
});

test("an async property", async () => {
  const seen: Kind[] = [];
  await assertProperty(
    fc.asyncProperty(fc.constantFrom(...KINDS), async (kind) => {
      seen.push(kind);
    }),
    { id: "async" },
  );
  expect(seen).toContain("a"); // caught
});

test("a sample that names no seed", () => {
  const seen = new Set<number>();
  for (const value of fc.sample(fc.nat(), 50)) seen.add(value % 3);
  expect(seen.size).toBe(3); // caught
});

test("a non-empty check over a list the predicate fills", () => {
  const reached: number[] = [];
  assertProperty(
    fc.property(fc.nat(), (value) => {
      reached.push(value);
    }),
    { id: "not empty" },
  );
  expect(reached).not.toEqual([]); // caught
});

const lawOver = (seen: Set<Kind>) =>
  fc.property(fc.constantFrom(...KINDS), (kind) => {
    seen.add(kind);
  });

test("a property a function builds over the tally", () => {
  const seen = new Set<Kind>();
  assertProperty(lawOver(seen), { id: "built" });
  expect(seen.size).toBe(KINDS.length); // caught
});

test("the same, through a binding", () => {
  const seen = new Set<Kind>();
  const property = lawOver(seen);
  assertProperty(property, { id: "bound" });
  expect(seen.has("a")).toBe(true); // caught
});

test("a sample that names no seed, stored before it is read", () => {
  const seen = new Set<number>();
  const values = fc.sample(fc.nat(), 50);
  for (const value of values) seen.add(value % 3);
  expect(seen.size).toBe(3); // caught
});

test("a collection the predicate writes through another name", () => {
  const seen = new Set<Kind>();
  const alias = seen;
  assertProperty(
    fc.property(fc.constantFrom(...KINDS), (kind) => {
      alias.add(kind);
    }),
    { id: "alias" },
  );
  expect(seen.size).toBeGreaterThan(0); // caught
});
