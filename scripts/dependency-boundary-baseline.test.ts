import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import path from "node:path";

const BASELINE_PATH = path.join(
  import.meta.dirname,
  "..",
  ".dependency-cruiser-known-violations.json",
);

describe("dependency boundary baseline", () => {
  test("does not permit package-boundary exceptions", async () => {
    const baseline: unknown = JSON.parse(await readFile(BASELINE_PATH, "utf8"));

    expect(baseline).toEqual([]);
  });
});
