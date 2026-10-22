import { expect, test } from "bun:test";

import {
  FAILURE_INVARIANTS,
  RELATION_INVARIANTS,
} from "../test/consumer-scenarios/support/failure-invariants";
import { RELATIONS } from "../test/consumer-scenarios/support/relation-contract";

test("every metamorphic relation has a stable failure invariant", () => {
  expect(Object.keys(RELATION_INVARIANTS).toSorted()).toEqual([...RELATIONS].toSorted());
  for (const invariant of Object.values(RELATION_INVARIANTS)) {
    expect(FAILURE_INVARIANTS[invariant].title.length).toBeGreaterThan(0);
  }
});

test("every oracle failure-producing comparison has an invariant", () => {
  expect(Object.keys(FAILURE_INVARIANTS).toSorted()).toEqual([
    "batchSequential",
    "blockPlacement",
    "commentAnchor",
    "commentLifetime",
    "directTracked",
    "numberingKind",
    "numberingLabel",
    "numberingLevel",
    "readerStability",
    "rejectAll",
    "rejectionRoundTrip",
    "requestedOutcome",
    "saveIdempotent",
    "undo",
  ]);
});
