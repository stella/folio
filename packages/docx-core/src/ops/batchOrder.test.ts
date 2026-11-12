import { describe, expect, test } from "bun:test";

import {
  deletionRemovesBlock,
  finalBreakCarrier,
  groupInsertionRuns,
  orderAddressingBeforeRemoval,
} from "./batchOrder";

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

describe("orderAddressingBeforeRemoval", () => {
  type Access = { id: string; addresses: string[]; removes: string[] };
  const order = (operations: Access[]) =>
    orderAddressingBeforeRemoval(operations, (operation) => operation).map(
      ({ id }) => id,
    );

  test("moves an operation on a block ahead of the one that removes it", () => {
    expect(
      order([
        { id: "delete-b", addresses: [], removes: ["b"] },
        { id: "edit-a", addresses: ["a"], removes: [] },
        { id: "edit-b", addresses: ["b"], removes: [] },
      ]),
    ).toEqual(["edit-b", "delete-b", "edit-a"]);
  });

  test("keeps every other operation in place", () => {
    const operations: Access[] = [
      { id: "edit-b", addresses: ["b"], removes: [] },
      { id: "merge-a", addresses: ["a"], removes: ["b"] },
      { id: "edit-c", addresses: ["c"], removes: [] },
    ];
    expect(order(operations)).toEqual(["edit-b", "merge-a", "edit-c"]);
  });

  test("puts a merge that addresses a deleted block before the deletion", () => {
    expect(
      order([
        { id: "delete-a", addresses: [], removes: ["a"] },
        { id: "merge-a", addresses: ["a"], removes: ["b"] },
        { id: "edit-b", addresses: ["b"], removes: [] },
      ]),
    ).toEqual(["edit-b", "merge-a", "delete-a"]);
  });
});

describe("deletionRemovesBlock", () => {
  test("only a container's last paragraph with nothing kept to take its break stays", () => {
    expect(
      deletionRemovesBlock({ endsContainer: true, carrier: "runStart" }),
    ).toBe(false);
    expect(deletionRemovesBlock({ endsContainer: true, carrier: "none" })).toBe(
      false,
    );
    expect(deletionRemovesBlock({ endsContainer: true, carrier: "kept" })).toBe(
      true,
    );
    expect(
      deletionRemovesBlock({ endsContainer: false, carrier: "none" }),
    ).toBe(true);
  });
});

describe("finalBreakCarrier", () => {
  const run = (deleted: readonly boolean[]) =>
    deleted.map((isDeleted, index) => ({
      paragraph: `p${index}`,
      deleted: isDeleted,
    }));

  test("the nearest kept paragraph takes the break, past deleted ones", () => {
    expect(finalBreakCarrier(run([true, true, false, true]))).toEqual({
      type: "kept",
      paragraph: "p2",
    });
  });

  test("a run deleted back to its start settles on the run's first paragraph", () => {
    for (let length = 1; length <= 5; length++) {
      expect(
        finalBreakCarrier(run(Array.from({ length }, () => true))),
      ).toEqual({ type: "runStart", paragraph: `p${length - 1}` });
    }
  });

  test("a paragraph with none before it has no carrier", () => {
    expect(finalBreakCarrier(run([]))).toEqual({ type: "none" });
  });
});
