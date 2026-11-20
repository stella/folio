import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  formatReplayCommand,
  parseReplayCommand,
  type ReplayCommand,
} from "../test/replay-command";

setDefaultTimeout(propertyTestTimeout(5_000));

const segment = fc.stringMatching(/^[\w@.-]{1,12}$/u);
// A digest names the case a path leads to: it is written only beside one.
const pathAndDigest = fc.oneof(
  fc.constant({ path: null }),
  fc.record(
    {
      path: fc
        .array(fc.nat({ max: 999 }), { minLength: 1, maxLength: 12 })
        .map((steps) => steps.join(":")),
      digest: fc.stringMatching(/^[0-9a-f]{8}$/u),
    },
    { requiredKeys: ["path"] },
  ),
);
// A line with no test file names no package and no test either.
const target = fc.oneof(
  fc.constant({ dir: "" }),
  fc.record(
    {
      dir: fc.oneof(
        fc.constant(""),
        segment.map((name) => `packages/${name}`),
      ),
      file: fc
        .array(segment, { minLength: 1, maxLength: 4 })
        .map((parts) => `${parts.join("/")}.test.ts`),
      pattern: fc.string({ maxLength: 40 }),
    },
    { requiredKeys: ["dir", "file"] },
  ),
);
const replayCommand: fc.Arbitrary<ReplayCommand> = fc
  .tuple(
    fc.record({
      seed: fc.integer(),
      factor: fc.constantFrom(1, 4, 10, 2.5),
    }),
    pathAndDigest,
    target,
  )
  .map(([run, replay, where]) => Object.assign({}, run, replay, where));

test("a replay line reads back as the replay it was written from", () => {
  assertProperty(
    fc.property(replayCommand, (command) => {
      expect(parseReplayCommand(formatReplayCommand(command))).toEqual(command);
    }),
    { id: "replay line round trip", numRuns: 300 },
  );
});

test("a line that is not a replay line reads as nothing", () => {
  expect(parseReplayCommand("bun test")).toBeUndefined();
  expect(
    parseReplayCommand("PROPERTY_TEST_SEED=1 bun test ./x.test.ts; rm -rf /"),
  ).toBeUndefined();
});
