import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-page-resources(release-before-page-removal)";

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
  return output.split(RULE_MARKER).length - 1;
};

describe("rendered page resource ownership", () => {
  test("rejects DOM removal without releasing resources first", () => {
    expect(lintFixture("page-resources.invalid.ts")).toBe(4);
  });

  test("accepts page removal after releasing resources", () => {
    expect(lintFixture("page-resources.valid.ts")).toBe(0);
  });
});
