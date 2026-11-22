import { expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { assertFreshIdentityEquivalent } from "./freshIdentityOracle";

const fixture = (ids: number[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: ids.map((id) => ({
            type: "run",
            content: [{ type: "text", text: "a" }],
            propertyChanges: [
              {
                type: "runPropertyChange",
                info: { id, author: "Earlier" },
                previousFormatting: { italic: false },
              },
            ],
          })),
        },
      ],
    },
  },
});

test("fresh identity equivalence preserves original slots and uses a single bijection", () => {
  const options = {
    original: fixture([7]),
    allocated: { revision: [2, 3, 4] },
  };
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      actual: fixture([7, 2]),
      expected: fixture([7, 3]),
    }),
  ).not.toThrow();
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      actual: fixture([7, 2, 2]),
      expected: fixture([7, 3, 4]),
    }),
  ).toThrow();
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      actual: fixture([7, 3, 2]),
      expected: fixture([7, 2, 3]),
    }),
  ).toThrow();
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      actual: fixture([7]),
      expected: fixture([2]),
    }),
  ).toThrow();
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      allocated: { revision: [7] },
      actual: fixture([7]),
      expected: fixture([7]),
    }),
  ).toThrow();
});

test("fresh identities follow allocation order within one operation of a plan", () => {
  const options = {
    original: fixture([7]),
    allocated: { revision: [2, 3, 4] },
    actual: fixture([7, 4, 2, 3]),
    expected: fixture([7, 2, 3, 4]),
  };
  const one = [new Set(["revision:2", "revision:3", "revision:4"])];
  const two = [new Set(["revision:2", "revision:3"]), new Set(["revision:4"])];
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      minted: { actual: two, expected: one },
    }),
  ).not.toThrow();
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      minted: { actual: one, expected: one },
    }),
  ).toThrow();
  expect(() =>
    assertFreshIdentityEquivalent({
      ...options,
      actual: fixture([7, 4, 3, 2]),
      minted: { actual: two, expected: one },
    }),
  ).toThrow();
});
