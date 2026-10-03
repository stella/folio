import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

import {
  areaOf,
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
  test("the workflow matrix exercises every selected file and collects every failure", () => {
    const workflow = Bun.YAML.parse(
      readFileSync(path.resolve(import.meta.dir, "../.github/workflows/ci.yml"), "utf8"),
    );
    const job = workflow.jobs["property-areas"];
    const files = propertyFiles();
    const exercised = [...job.strategy.matrix.shard].flatMap((index: number) =>
      shardPropertyFiles(files, { index, total: job.strategy.matrix.shard.length }),
    );
    expect(exercised.length).toBe(files.length);
    expect(new Set(exercised).size).toBe(files.length);
    expect(exercised.toSorted((a, b) => a.file.localeCompare(b.file))).toEqual(files);
    expect(job.strategy["fail-fast"]).toBe(false);
    const step = job.steps.find(
      (entry: { name?: string }) => entry.name === "Changed areas' properties at 5x numRuns",
    );
    expect(step.run).toContain('--factor 5 --shard "${PROPERTY_SHARD}/${PROPERTY_SHARD_COUNT}"');
    expect(step.env.PROPERTY_SHARD).toBe("${{ matrix.shard }}");
  });

  test("shards cover every selected file exactly once regardless of input order", () => {
    const inventory = propertyFiles();
    for (const total of [1, 2, 4, inventory.length + 1]) {
      const exercised: string[] = [];
      for (let index = 1; index <= total; index += 1) {
        const shard = { index, total };
        const files = shardPropertyFiles(inventory, shard);
        expect(shardPropertyFiles(inventory.toReversed(), shard)).toEqual(files);
        for (const { file } of files) exercised.push(file);
      }
      expect(exercised.toSorted()).toEqual(inventory.map(({ file }) => file).toSorted());
      expect(new Set(exercised).size).toBe(exercised.length);
    }
    expect(shardPropertyFiles([], { index: 1, total: 4 })).toEqual([]);
  });

  test("invalid shards fail rather than silently skipping properties", () => {
    for (const shard of [
      { index: 0, total: 4 },
      { index: 5, total: 4 },
      { index: 1, total: 0 },
      { index: 1.5, total: 4 },
      { index: 1, total: Number.NaN },
    ])
      expect(() => shardPropertyFiles(ALL, shard)).toThrow();
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
