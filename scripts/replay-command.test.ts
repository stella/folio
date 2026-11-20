import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty } from "../test/property-testing";
import {
  formatReplayCommand,
  parseReplayCommand,
  type ReplayCommand,
} from "../test/replay-command";

const segment = fc.stringMatching(/^[\w@.-]{1,12}$/u);
const replayCommand: fc.Arbitrary<ReplayCommand> = fc
  .record(
    {
      dir: fc.oneof(
        fc.constant(""),
        segment.map((name) => `packages/${name}`),
      ),
      file: fc
        .array(segment, { minLength: 1, maxLength: 4 })
        .map((parts) => `${parts.join("/")}.test.ts`),
      pattern: fc.string({ maxLength: 40 }),
      seed: fc.integer(),
      path: fc.option(
        fc
          .array(fc.nat({ max: 999 }), { minLength: 1, maxLength: 12 })
          .map((steps) => steps.join(":")),
        { nil: null },
      ),
      digest: fc.stringMatching(/^[0-9a-f]{8}$/u),
      factor: fc.constantFrom(1, 4, 10, 2.5),
    },
    { requiredKeys: ["dir", "seed", "path", "factor"] },
  )
  // A digest names the case a path leads to: it is written only beside one.
  .map((command) => (command.path === null ? omitDigest(command) : command))
  // A line with no test file names no package either.
  .map((command) =>
    command.file === undefined ? { ...omitPattern(command), dir: "" } : command,
  );

const omitDigest = ({ digest: _digest, ...rest }: ReplayCommand) => rest;
const omitPattern = ({ pattern: _pattern, ...rest }: ReplayCommand) => rest;

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
