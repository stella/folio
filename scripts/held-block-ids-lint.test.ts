import { expect, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const ROOT = path.resolve(import.meta.dir, "..");
const lintFixture = (kind: "invalid" | "valid") => {
  const result = expectExited(
    runChild(
      [
        "bun",
        "--bun",
        "oxlint",
        "-c",
        "oxlint.fixtures.config.ts",
        "--no-ignore",
        path.join("test/__fixtures__/held-block-ids", `${kind}.ts`),
      ],
      { cwd: ROOT },
    ),
  );
  const output = `${result.stdout}${result.stderr}`;
  expect(output).not.toContain("Failed to load");
  expect(output).not.toContain("Failed to parse");
  return (
    output.split("folio-held-block-ids(reread-after-resolution)").length - 1
  );
};

test(
  "a block id held across accepting or rejecting is read again before use",
  () => {
    expect(lintFixture("invalid")).toBe(7);
    expect(lintFixture("valid")).toBe(0);
  },
  subprocessTestTimeout(2),
);
