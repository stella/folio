import { describe, expect, test } from "bun:test";

import { areaOf, propertyFiles, selectPropertyFiles, touchedAreas } from "./property-areas";

const file = (repoPath: string) => ({
  area: areaOf(repoPath) as string,
  packageDir: repoPath.split("/").slice(0, 2).join("/"),
  file: repoPath,
});

const ALL = [
  file("packages/core/src/ai-edits/batchOverlap.test.ts"),
  file("packages/core/src/compare/compare.property.test.ts"),
  file("packages/core/src/docx/xmlSerialize.property.test.ts"),
  file("packages/docx-core/src/markdown/markdown.test.ts"),
];

describe("property areas", () => {
  test("an area is the package and the first directory under src", () => {
    expect(areaOf("packages/core/src/compare/diff/align.ts")).toBe("core/compare");
    expect(areaOf("packages/docx-core/src/index.ts")).toBe("docx-core");
    expect(areaOf("scripts/property-areas.ts")).toBeUndefined();
  });

  test("changed source selects its area and the areas coupled to it", () => {
    expect([...touchedAreas(["packages/core/src/internal/compare/blocks.ts"])].toSorted()).toEqual([
      "core/ai-edits",
      "core/compare",
      "core/internal",
      "core/prosemirror",
    ]);
    expect(
      selectPropertyFiles(["packages/docx-core/src/markdown/compile.ts"], ALL).map((f) => f.file),
    ).toEqual(["packages/docx-core/src/markdown/markdown.test.ts"]);
  });

  test("a changed test file reruns itself, not its area", () => {
    expect(
      selectPropertyFiles(
        ["packages/core/src/compare/compare.property.test.ts", "packages/core/src/docx/a.test.ts"],
        ALL,
      ).map((f) => f.file),
    ).toEqual(["packages/core/src/compare/compare.property.test.ts"]);
  });

  test("a change to the pinned seeds reruns the files they name", () => {
    expect(
      selectPropertyFiles(["test/property-seeds.json"], ALL, [
        "packages/core/src/ai-edits/batchOverlap.test.ts",
      ]).map((f) => f.file),
    ).toEqual(["packages/core/src/ai-edits/batchOverlap.test.ts"]);
    expect(selectPropertyFiles(["README.md"], ALL)).toEqual([]);
  });

  test("finds the property files of every package", () => {
    const found = propertyFiles();
    expect(found.some(({ file: path }) => path.startsWith("packages/core/src/"))).toBe(true);
    expect(found.some(({ file: path }) => path.startsWith("packages/docx-core/src/"))).toBe(true);
    expect(found.every(({ area }) => area !== undefined)).toBe(true);
  });
});
