import { describe, expect, test } from "bun:test";

import {
  firstBuildTreeMismatch,
  renderBuildTreeMismatch,
  type BuildTree,
} from "./lib/build-tree-diff";

const tree = (files: Record<string, string>): BuildTree => new Map(Object.entries(files));

const mismatchOf = (
  baseline: Record<string, string>,
  candidate: Record<string, string>,
): ReturnType<typeof firstBuildTreeMismatch> =>
  firstBuildTreeMismatch({ baseline: tree(baseline), candidate: tree(candidate) });

describe("firstBuildTreeMismatch", () => {
  test("accepts two builds that emitted the same files", () => {
    expect(
      mismatchOf({ "index.js": "a", "index.d.ts": "b" }, { "index.d.ts": "b", "index.js": "a" }),
    ).toBeNull();
  });

  test("accepts an empty pair of trees", () => {
    expect(mismatchOf({}, {})).toBeNull();
  });

  test("names the differing file and carries both versions", () => {
    expect(mismatchOf({ "index.js": "one" }, { "index.js": "two" })).toEqual({
      kind: "content",
      path: "index.js",
      baseline: "one",
      candidate: "two",
    });
  });

  test("reports a declaration difference, not only JavaScript", () => {
    const mismatch = mismatchOf({ "types.d.ts": "type A = 1;" }, { "types.d.ts": "type A = 2;" });
    expect(mismatch?.kind).toBe("content");
    expect(mismatch?.path).toBe("types.d.ts");
  });

  test("picks the first path in sorted order when several differ", () => {
    expect(
      mismatchOf(
        { "z.js": "1", "a.js": "1", "m.js": "1" },
        { "z.js": "2", "a.js": "2", "m.js": "2" },
      )?.path,
    ).toBe("a.js");
  });

  test("sorts by path, not by the order either tree was walked", () => {
    expect(mismatchOf({ "b.js": "x", "a.js": "1" }, { "b.js": "y", "a.js": "2" })?.path).toBe(
      "a.js",
    );
  });

  test("reports a file only the first build emitted", () => {
    expect(mismatchOf({ "chunk-AAA.js": "x" }, {})).toEqual({
      kind: "only-in-baseline",
      path: "chunk-AAA.js",
    });
  });

  test("reports a file only the second build emitted", () => {
    expect(mismatchOf({}, { "chunk-BBB.js": "x" })).toEqual({
      kind: "only-in-candidate",
      path: "chunk-BBB.js",
    });
  });
});

describe("renderBuildTreeMismatch", () => {
  const render = (
    baseline: Record<string, string>,
    candidate: Record<string, string>,
    maxDiffLines = 200,
  ): string => {
    const mismatch = mismatchOf(baseline, candidate);
    if (!mismatch) throw new Error("expected a mismatch");
    return renderBuildTreeMismatch({ mismatch, maxDiffLines });
  };

  test("prints the path, then the diff of its two versions", () => {
    expect(
      render({ "index.js": "export const a = 1;" }, { "index.js": "export const a = 2;" }),
    ).toBe("index.js\n-export const a = 1;\n+export const a = 2;");
  });

  test("caps the diff and says how many lines it dropped", () => {
    const baseline = Array.from({ length: 30 }, (_, index) => `old ${index}`).join("\n");
    const candidate = Array.from({ length: 30 }, (_, index) => `new ${index}`).join("\n");
    const lines = render({ "index.d.ts": baseline }, { "index.d.ts": candidate }, 6).split("\n");

    expect(lines.at(0)).toBe("index.d.ts");
    expect(lines).toHaveLength(8);
    expect(lines.at(-1)).toBe("… 54 more diff line(s) truncated");
  });

  test("does not truncate a difference that fits", () => {
    expect(render({ "a.js": "x" }, { "a.js": "y" }, 2)).toBe("a.js\n-x\n+y");
  });

  test("says which build emitted a file the other did not", () => {
    expect(render({ "chunk-AAA.js": "x" }, {})).toBe(
      "chunk-AAA.js\n  emitted by the first build only",
    );
    expect(render({}, { "chunk-BBB.js": "x" })).toBe(
      "chunk-BBB.js\n  emitted by the second build only",
    );
  });
});
