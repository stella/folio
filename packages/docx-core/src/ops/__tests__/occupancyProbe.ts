import { expect } from "bun:test";

import type { PackageIdOccupancy } from "../editorIntentCensus";

/** The ids a census read afresh holds, by space. */
export type CensusIds = {
  revisions: ReadonlySet<number>;
  controls: ReadonlySet<number>;
  paragraphs: ReadonlySet<string>;
};

/** Ids beside each held one, and the low ids an allocation tries first. */
const probes = (held: ReadonlySet<number>): number[] => [
  ...new Set([
    ...Array.from({ length: 12 }, (_, id) => id),
    ...[...held].flatMap((id) => [id - 1, id, id + 1]),
  ]),
];

/**
 * Require `occupancy` to answer as the census does: every id it holds is
 * taken, and the ids beside them and the lowest ids are taken only when it
 * holds them.
 */
export const expectOccupancyOf = (
  occupancy: PackageIdOccupancy,
  census: CensusIds,
): void => {
  const wrong = [
    ...probes(census.revisions)
      .filter((id) => occupancy.revisions.has(id) !== census.revisions.has(id))
      .map((id) => `revision ${String(id)}`),
    ...probes(census.controls)
      .filter((id) => occupancy.controls.has(id) !== census.controls.has(id))
      .map((id) => `control ${String(id)}`),
    ...[...census.paragraphs, "00000000", "00000001", "7FFFFFFF"]
      .filter(
        (key) => occupancy.paragraphs.has(key) !== census.paragraphs.has(key),
      )
      .map((key) => `paragraph ${key}`),
  ];
  expect(wrong).toEqual([]);
};
