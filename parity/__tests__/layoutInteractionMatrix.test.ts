import { describe, expect, test } from "bun:test";

import {
  buildLayoutInteractionMatrix,
  layoutInteractionCasePairs,
  LAYOUT_INTERACTION_AXES,
  LAYOUT_INTERACTION_AXIS_NAMES,
  validLayoutInteractionPairs,
} from "../fixtures/layout-interaction-matrix";

describe("synthetic layout interaction matrix", () => {
  test("is deterministic, compact, and content-addressed", () => {
    const first = buildLayoutInteractionMatrix();
    const second = buildLayoutInteractionMatrix();
    const cartesianSize = Object.values(LAYOUT_INTERACTION_AXES).reduce(
      (product, values) => product * values.length,
      1,
    );

    expect(first).toEqual(second);
    expect(new Set(first.map(({ id }) => id)).size).toBe(first.length);
    expect(first.every(({ id }) => /^mx-[0-9a-f]{10}$/u.test(id))).toBeTrue();
    expect(first.length).toBeLessThan(cartesianSize / 100);
  });

  test("covers every pair that belongs to a valid complete scenario", () => {
    const covered = new Set<string>();
    for (const scenario of buildLayoutInteractionMatrix()) {
      for (const pair of layoutInteractionCasePairs(scenario)) covered.add(pair);
    }

    expect([...covered].toSorted()).toEqual([...validLayoutInteractionPairs()].toSorted());
  });

  test("uses every declared axis value", () => {
    const matrix = buildLayoutInteractionMatrix();

    for (const axis of LAYOUT_INTERACTION_AXIS_NAMES) {
      expect(new Set(matrix.map((scenario) => scenario[axis]))).toEqual(
        new Set(LAYOUT_INTERACTION_AXES[axis]),
      );
    }
  });
});
