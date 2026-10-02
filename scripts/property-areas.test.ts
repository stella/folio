import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

import {
  areaOf,
  parsePropertyShard,
  propertyFiles,
  selectPropertyFiles,
  shardPropertyFiles,
  touchedAreas,
} from "./property-areas";

const makePropertyFile = (repoPath: string) => ({
  area: areaOf(repoPath) as string,
  packageDir: repoPath.split("/").slice(0, 2).join("/"),
  file: repoPath,
});

const ALL = [
  makePropertyFile("packages/core/src/ai-edits/batchOverlap.test.ts"),
  makePropertyFile("packages/core/src/compare/compare.property.test.ts"),
  makePropertyFile("packages/core/src/docx/xmlSerialize.property.test.ts"),
  makePropertyFile("packages/docx-core/src/markdown/markdown.test.ts"),
];

/** Independently inventory actual calls, excluding matches in comments and strings. */
const actualPropertyDrivers = (): string[] => {
  const root = path.resolve(import.meta.dir, "..");
  const files = ["packages", "scripts", "test", "parity"].flatMap((dir) =>
    ts.sys.readDirectory(path.join(root, dir), [".ts", ".tsx"], ["**/node_modules/**"]),
  );
  return files
    .filter((file) => /\.test\.tsx?$/.test(file))
    .filter((file) => {
      const source = ts.createSourceFile(
        file,
        readFileSync(file, "utf8"),
        ts.ScriptTarget.Latest,
        true,
      );
      let drives = false;
      const visit = (node: ts.Node): void => {
        if (ts.isCallExpression(node)) {
          const callee = node.expression;
          drives ||=
            (ts.isIdentifier(callee) && callee.text === "assertProperty") ||
            (ts.isPropertyAccessExpression(callee) &&
              ts.isIdentifier(callee.expression) &&
              callee.expression.text === "fc" &&
              (callee.name.text === "assert" || callee.name.text === "check"));
        }
        if (!drives) ts.forEachChild(node, visit);
      };
      visit(source);
      return drives;
    })
    .map((file) => path.relative(root, file).replaceAll("\\", "/"))
    .toSorted((a, b) => a.localeCompare(b));
};

describe("property areas", () => {
  test("shards partition the actual selected files exactly once without changing package ownership", () => {
    const files = propertyFiles();
    for (const count of [1, 2, 4, files.length + 1]) {
      const shards = Array.from({ length: count }, (_, index) =>
        shardPropertyFiles(files, parsePropertyShard(`${String(index + 1)}/${String(count)}`)),
      );
      const partitioned = shards.flat();
      expect(partitioned.length).toBe(files.length);
      expect(new Set(partitioned).size).toBe(files.length);
      expect(partitioned.toSorted((a, b) => a.file.localeCompare(b.file))).toEqual(files);
      expect(
        Math.max(...shards.map((shard) => shard.length)) -
          Math.min(...shards.map((shard) => shard.length)),
      ).toBeLessThanOrEqual(1);
    }
  });

  test("the workflow matrix exercises every selected file and collects every failure", () => {
    const workflow = Bun.YAML.parse(
      readFileSync(path.resolve(import.meta.dir, "../.github/workflows/ci.yml"), "utf8"),
    );
    const job = workflow.jobs["property-areas"];
    const files = propertyFiles();
    const exercised = [...job.strategy.matrix.shard].flatMap((raw: string) =>
      shardPropertyFiles(files, parsePropertyShard(raw)),
    );
    expect(exercised.length).toBe(files.length);
    expect(new Set(exercised).size).toBe(files.length);
    expect(exercised.toSorted((a, b) => a.file.localeCompare(b.file))).toEqual(files);
    expect(job.strategy["fail-fast"]).toBe(false);
    const step = job.steps.find(
      (entry: { name?: string }) => entry.name === "Changed areas' properties at 5x numRuns",
    );
    expect(step.run).toContain('--factor 5 --shard "${PROPERTY_SHARD}"');
    expect(step.env.PROPERTY_SHARD).toBe("${{ matrix.shard }}");
  });

  test("invalid shard coordinates fail instead of silently dropping coverage", () => {
    for (const raw of [
      "",
      "0/4",
      "5/4",
      "1/0",
      "1",
      "1/2/3",
      "-1/4",
      "1.5/4",
      "1/9007199254740992",
    ])
      expect(() => parsePropertyShard(raw)).toThrow();
  });

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
    expect(found.some(({ file: repoPath }) => repoPath.startsWith("packages/core/src/"))).toBe(
      true,
    );
    expect(found.some(({ file: repoPath }) => repoPath.startsWith("packages/docx-core/src/"))).toBe(
      true,
    );
    expect(found.every(({ area }) => area !== undefined)).toBe(true);
    expect(
      found.some(
        ({ file: repoPath }) => repoPath === "scripts/container-survival.property.test.ts",
      ),
    ).toBe(true);
    expect(found.map(({ file }) => file)).toEqual(actualPropertyDrivers());
  });

  test("a root script change selects its properties", () => {
    const rootProperty = {
      area: "scripts",
      packageDir: ".",
      file: "scripts/container-survival.property.test.ts",
    };
    expect(selectPropertyFiles(["scripts/container-survival-census.ts"], [rootProperty])).toEqual([
      rootProperty,
    ]);
  });
});
