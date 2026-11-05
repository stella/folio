import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

/** Scans the sources so a new inline node cannot miss this by not being listed. */
const inlineExtensions = readdirSync(import.meta.dir).filter(
  (file) =>
    file.endsWith("Extension.ts") &&
    readFileSync(join(import.meta.dir, file), "utf8").includes("inline: true"),
);

describe("inline node DOM", () => {
  test("the scan sees the inline anchors", () => {
    expect(inlineExtensions).toContain("BookmarkBoundaryExtension.ts");
    expect(inlineExtensions).toContain("CommentRangeBoundaryExtension.ts");
  });

  for (const file of inlineExtensions) {
    test(`${file} keeps a box in the editor DOM`, () => {
      expect(readFileSync(join(import.meta.dir, file), "utf8")).not.toMatch(
        /display:\s*none/u,
      );
    });
  }
});
