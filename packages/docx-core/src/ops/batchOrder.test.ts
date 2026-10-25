import { describe, expect, test } from "bun:test";

import { groupInsertionRuns } from "./batchOrder";

type Operation = { id: string; gap: number | null; writesProperties?: true };

const ids = (groups: Operation[][]) =>
  groups.map((group) => group.map(({ id }) => id));

describe("groupInsertionRuns", () => {
  const gapOf = ({ gap }: Operation) => gap;
  const keepsGap = ({ writesProperties }: Operation) =>
    writesProperties === true;

  test("joins consecutive insertions at one gap", () => {
    expect(
      ids(
        groupInsertionRuns(
          [
            { id: "a", gap: 5 },
            { id: "b", gap: 5 },
            { id: "c", gap: 2 },
          ],
          gapOf,
          keepsGap,
        ),
      ),
    ).toEqual([["a", "b"], ["c"]]);
  });

  test("keeps a run open across a property write at its gap (#2158)", () => {
    expect(
      ids(
        groupInsertionRuns(
          [
            { id: "a", gap: 5 },
            { id: "props", gap: null, writesProperties: true },
            { id: "b", gap: 5 },
          ],
          gapOf,
          keepsGap,
        ),
      ),
    ).toEqual([["a", "b"], ["props"]]);
  });

  test("ends a run at any other operation", () => {
    expect(
      ids(
        groupInsertionRuns(
          [
            { id: "a", gap: 5 },
            { id: "merge", gap: null },
            { id: "b", gap: 5 },
          ],
          gapOf,
          keepsGap,
        ),
      ),
    ).toEqual([["a"], ["merge"], ["b"]]);
  });

  test("without a predicate, every other operation ends a run", () => {
    expect(
      ids(
        groupInsertionRuns(
          [
            { id: "a", gap: 5 },
            { id: "props", gap: null, writesProperties: true },
            { id: "b", gap: 5 },
          ],
          gapOf,
        ),
      ),
    ).toEqual([["a"], ["props"], ["b"]]);
  });
});
