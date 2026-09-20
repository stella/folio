/** Wiring test for the xml-splice-owner lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-xml-splice(no-hand-rolled-splice)";

setDefaultTimeout(30_000);

const lintFixture = (fixture: string) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.config.ts",
      "--no-ignore",
      path.join("test", "__fixtures__", fixture),
    ],
    { cwd: REPO_ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  return output.split(RULE_MARKER).length - 1;
};

describe("xml splice owner", () => {
  test("rejects every spelling of a region cut out of a part", () => {
    expect(lintFixture("xml-splice-owner.invalid.ts")).toBe(3);
  });

  test("accepts the owner, an append, and slices that meet", () => {
    expect(lintFixture("xml-splice-owner.valid.ts")).toBe(0);
  });
});
