import { expect, test } from "bun:test";
import path from "node:path";

const ROOT = path.resolve(import.meta.dir, "..");
const lintFixture = (kind: string) => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.fixtures.config.ts",
      "--no-ignore",
      path.join("test/__fixtures__/deletion-marks", kind, "packages/core/src/mutations.ts"),
    ],
    { cwd: ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  expect(output).not.toContain("Failed to load");
  expect(output).not.toContain("Failed to parse");
  return output.split("folio-deletion-marks(preserve-pending-deletions)").length - 1;
};

test("fresh deletion marks use the owner through aliases and schema constructors", () => {
  expect(lintFixture("invalid")).toBe(9);
  expect(lintFixture("valid")).toBe(0);
}, 30_000);

test("deletion aliases resolve their lexical binding across siblings and shadowing", () => {
  expect(lintFixture("scoped")).toBe(5);
  expect(lintFixture("scoped-valid")).toBe(0);
}, 30_000);
