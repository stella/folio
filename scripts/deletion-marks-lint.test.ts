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
      "oxlint.config.ts",
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
  expect(lintFixture("invalid")).toBe(7);
  expect(lintFixture("valid")).toBe(0);
}, 30_000);
