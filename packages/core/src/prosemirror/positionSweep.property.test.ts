/**
 * `sweepPositions` answers exactly what the per-step walk answers: for every
 * query, `mapping.slice(from).mapResult(pos, assoc)` — the position and
 * whether a step deleted across it — whatever mix of steps the transaction
 * holds, wherever the query enters, and however the steps are split between
 * transactions.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { Mapping, StepMap } from "prosemirror-transform";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  buildRandomTransaction,
  randomDocument,
  randomOperations,
} from "./__tests__/randomTransactions";
import { sweepPositions, type PositionQuery } from "./positionSweep";

setDefaultTimeout(propertyTestTimeout(30_000));

const perStep = (maps: readonly StepMap[], queries: readonly PositionQuery[]) =>
  queries.map(({ pos, assoc, from }) => {
    const result = new Mapping(maps).slice(from).mapResult(pos, assoc);
    return { pos: result.pos, deletedAcross: result.deletedAcross };
  });

const rawQuery = fc.record({
  pos: fc.nat({ max: 10_000 }),
  assoc: fc.constantFrom(-1 as const, 1 as const),
  from: fc.nat({ max: 10_000 }),
});

describe("sweepPositions", () => {
  test("maps every query to the per-step result", () => {
    fc.assert(
      fc.property(
        randomDocument,
        randomOperations(30),
        fc.array(rawQuery, { maxLength: 400, size: "large" }),
        fc.array(fc.nat({ max: 10_000 }), { maxLength: 3 }),
        (doc, operations, rawQueries, cuts) => {
          const tr = buildRandomTransaction(EditorState.create({ doc }), operations);
          const maps = tr.mapping.maps;
          const docs = [...tr.docs, tr.doc];
          const queries = rawQueries.map(({ pos, assoc, from }): PositionQuery => {
            const step = from % (maps.length + 1);
            // SAFETY: docs holds one document per step plus the final one.
            const size = docs[step]!.content.size;
            return { pos: pos % (size + 1), assoc, from: step };
          });
          // The same steps split into consecutive transactions' mappings.
          const boundaries = [
            0,
            ...cuts.map((cut) => cut % (maps.length + 1)).sort((a, b) => a - b),
            maps.length,
          ];
          const mappings = boundaries
            .slice(1)
            .map((end, index) => new Mapping(maps.slice(boundaries[index], end)));

          expect(sweepPositions(mappings, queries)).toEqual(perStep(maps, queries));
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("keeps positions ordered when many enter at one step and chunks split", () => {
    // One insertion at the front of a long range shifts everything after it;
    // a deletion then collapses a run of positions onto one boundary.
    const maps = [
      new StepMap([0, 0, 5]),
      new StepMap([100, 400, 0]),
      new StepMap([50, 10, 20, 300, 0, 7]),
    ];
    const queries: PositionQuery[] = [];
    for (let pos = 0; pos <= 1_000; pos += 1) {
      queries.push({ pos, assoc: pos % 2 === 0 ? 1 : -1, from: pos % 3 === 0 ? 0 : 1 });
    }
    for (let pos = 0; pos <= 600; pos += 1) {
      queries.push({ pos, assoc: 1, from: 2 });
    }
    expect(sweepPositions([new Mapping(maps)], queries)).toEqual(perStep(maps, queries));
  });

  test("falls back to the per-step walk when a mapping carries mirrors", () => {
    // A deletion followed by its own inverse, mirrored: the deleted interior
    // is recovered rather than collapsed onto the deletion's edge.
    const deletion = new StepMap([5, 10, 0]);
    const mapping = new Mapping();
    mapping.appendMap(deletion);
    mapping.appendMap(deletion.invert(), 0);
    const queries: PositionQuery[] = [7, 9, 12, 20].map((pos) => ({ pos, assoc: 1, from: 0 }));

    expect(sweepPositions([mapping], queries)).toEqual(
      queries.map(({ pos, assoc }) => {
        const result = mapping.mapResult(pos, assoc);
        return { pos: result.pos, deletedAcross: result.deletedAcross };
      }),
    );
    // Recovered, not collapsed: the plain per-step walk would say 5 or 15.
    expect(sweepPositions([mapping], queries).map(({ pos }) => pos)).toEqual([7, 9, 12, 20]);
  });
});
