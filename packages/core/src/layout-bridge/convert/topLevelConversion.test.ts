import { describe, expect, test } from "bun:test";

import type { ListCounterState } from "../../prosemirror/listMarker";
import {
  restoreListCounterState,
  sameListCounterState,
} from "./topLevelConversion";

/** Numbering instance 1 and its abstract definition 7 advancing one array. */
const sharedState = (): ListCounterState => {
  const shared = [2, Number.NaN, Number.NaN];
  return {
    counters: new Map([
      [1, shared],
      [2, [5, 1, Number.NaN]],
    ]),
    abstractCounters: new Map([[7, shared]]),
    seenLevels: new Set(["1:0", "2:0", "2:1"]),
    restartedNumIds: new Set([2]),
    previousList: { abstractNumId: 7, fromStyle: false, numId: 1 },
    lastAdvancedNumId: 1,
  };
};

const emptyState = (): ListCounterState => ({
  counters: new Map(),
  abstractCounters: new Map(),
  seenLevels: new Set(),
  restartedNumIds: new Set(),
  previousList: { abstractNumId: null, fromStyle: false, numId: null },
});

describe("top-level conversion state", () => {
  test("restoring a state keeps the arrays it shares shared, in arrays of its own", () => {
    const source = sharedState();
    const target = emptyState();
    const { counters, abstractCounters, seenLevels } = target;
    restoreListCounterState(target, source);

    expect(sameListCounterState(target, source)).toBe(true);
    // The maps and sets the conversion options hold are the ones restored.
    expect(target.counters).toBe(counters);
    expect(target.abstractCounters).toBe(abstractCounters);
    expect(target.seenLevels).toBe(seenLevels);
    expect(target.counters.get(1)).toBe(target.abstractCounters.get(7));
    expect(target.counters.get(1)).not.toBe(source.counters.get(1));

    // Advancing the restored state leaves the checkpoint as it was.
    target.counters.get(1)![0] = 3;
    expect(source.counters.get(1)![0]).toBe(2);
    expect(target.abstractCounters.get(7)![0]).toBe(3);
  });

  test("two states differ when only which arrays they share differs", () => {
    const shared = sharedState();
    const apart = sharedState();
    apart.abstractCounters.set(7, [...apart.abstractCounters.get(7)!]);
    expect(sameListCounterState(shared, apart)).toBe(false);
    expect(sameListCounterState(shared, sharedState())).toBe(true);
  });

  test("two states differ in a counter that has none yet", () => {
    const left = sharedState();
    const right = sharedState();
    right.counters.get(2)![2] = 0;
    expect(sameListCounterState(left, right)).toBe(false);
  });
});
