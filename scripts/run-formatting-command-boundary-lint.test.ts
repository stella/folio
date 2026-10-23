import { expect, test } from "bun:test";
import path from "node:path";

const ROOT = path.resolve(import.meta.dir, "..");
const MARKER = "folio-editor-commands(run-formatting-lock-boundary)";
const lintFixture = (kind: "invalid" | "valid") => {
  const result = Bun.spawnSync(
    [
      "bun",
      "--bun",
      "oxlint",
      "-c",
      "oxlint.fixtures.config.ts",
      "--no-ignore",
      path.join(
        "test/__fixtures__/run-formatting-command-boundary",
        kind,
        "packages/core/src/prosemirror/extensions/marks/BoldExtension.ts",
      ),
    ],
    { cwd: ROOT },
  );
  const output = `${result.stdout.toString()}${result.stderr.toString()}`;
  expect(output).not.toContain("Failed to load");
  expect(output).not.toContain("Failed to parse");
  return output.split(MARKER).length - 1;
};

test("run-formatting extension commands use the content-lock boundary", () => {
  expect(lintFixture("invalid")).toBe(1);
  expect(lintFixture("valid")).toBe(0);
}, 30_000);
