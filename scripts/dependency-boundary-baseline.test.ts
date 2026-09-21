import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import path from "node:path";

const CYCLE_RULE_NAME = "no-circular";
const BASELINE_PATH = path.join(
  import.meta.dirname,
  "..",
  ".dependency-cruiser-known-violations.json",
);

describe("dependency boundary baseline", () => {
  test("does not permit package-boundary exceptions", async () => {
    const baseline: unknown = JSON.parse(await readFile(BASELINE_PATH, "utf8"));
    const packageBoundaryExceptions = Array.isArray(baseline)
      ? baseline.filter((entry) => {
          if (typeof entry !== "object" || entry === null || !("rule" in entry)) {
            return true;
          }
          const { rule } = entry;
          return (
            typeof rule !== "object" ||
            rule === null ||
            !("name" in rule) ||
            rule.name !== CYCLE_RULE_NAME
          );
        })
      : [baseline];

    expect(packageBoundaryExceptions).toEqual([]);
  });
});
