/** Wiring test for the model-read lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-read-model(no-model-document-cast)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.fixtures.config.ts",
      "--no-ignore",
      path.join("test", "__fixtures__", fixture),
    ],
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  expect(output).not.toContain("Failed to load");
  expect(output).not.toContain("Failed to parse");
  return output.split(RULE_MARKER).length - 1;
};

describe("model read owner", () => {
  test("rejects every cast of a model to the model Document", () => {
    expect(lintFixture("read-model-cast.invalid.ts")).toBe(8);
  });

  test("reads a single-file component's script block", () => {
    expect(lintFixture("read-model-cast.invalid.vue")).toBe(2);
  });

  test("accepts readModel, checked literals and the DOM Document", () => {
    expect(lintFixture("read-model-cast.valid.ts")).toBe(0);
  });
});
