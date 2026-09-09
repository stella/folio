import { describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { buildImportGraph, buildReport, checksForFile, classifySeam, resolveFiles } from "./orient";

const REPO_ROOT = path.resolve(import.meta.dir, "..");

const trackedFiles = (): readonly string[] => {
  const result = Bun.spawnSync(["git", "ls-files"], { cwd: REPO_ROOT, stdout: "pipe" });
  if (result.exitCode !== 0) {
    throw new Error("git ls-files failed in orientation test");
  }
  return result.stdout.toString().trim().split("\n");
};

describe("repository orientation", () => {
  test("classifies each fidelity pipeline seam before generic test files", () => {
    expect(classifySeam("packages/core/src/docx/parser.ts")).toBe("parser-serializer");
    expect(classifySeam("packages/core/src/layout-bridge/convert/toFlowBlocks.ts")).toBe(
      "flow-conversion",
    );
    expect(classifySeam("packages/core/src/layout-engine/measure/measureBlocks.ts")).toBe(
      "measurement-shaping",
    );
    expect(classifySeam("packages/core/src/layout-engine/paginator.ts")).toBe("pagination");
    expect(classifySeam("packages/core/src/display-list/build/buildDisplayList.ts")).toBe(
      "display-list",
    );
    expect(classifySeam("packages/core/src/display-list/dom/renderDisplayListToDom.ts")).toBe(
      "dom-render",
    );
    expect(classifySeam("packages/core/src/layout-engine/paginator.test.ts")).toBe("test-harness");
  });

  test("layout source changes select local and interaction validation", () => {
    expect(checksForFile("packages/core/src/layout-engine/paginator.ts")).toEqual([
      "bun --filter @stll/folio-core test",
      "bun --filter @stll/folio-core typecheck",
      "bun run test:interactions",
      "bun run test:measure-parity",
      "bun run lint",
      "bun run format:check",
    ]);
  });

  test("a real source file resolves ownership, instructions, tests, and release impact", () => {
    const file = "packages/core/src/docx/ensureParaIds.ts";
    const report = buildReport(REPO_ROOT, trackedFiles(), [file]);
    const orientation = report.files.at(0);

    expect(orientation?.package?.name).toBe("@stll/folio-core");
    expect(orientation?.seam).toBe("parser-serializer");
    expect(orientation?.instructions).toEqual(["AGENTS.md", "packages/core/AGENTS.md"]);
    expect(orientation?.tests).toContain("packages/core/src/docx/ensureParaIds.test.ts");
    expect(orientation?.importedBy).toContain("packages/core/src/docx/ensureParaIds.test.ts");
    expect(orientation?.requiresChangeset).toBe(true);
    expect(report.changesetPackages).toEqual(["@stll/folio-core"]);
  });

  test("finds tests through a source entry point", () => {
    const report = buildReport(REPO_ROOT, trackedFiles(), [
      "packages/core/src/display-list/build/tablePrimitives.ts",
    ]);

    expect(report.files.at(0)?.tests).toContain(
      "packages/core/src/display-list/clippedTableRegions.test.ts",
    );
  });

  test("expands directories deterministically and removes duplicate files", () => {
    const files = resolveFiles({
      repoRoot: REPO_ROOT,
      trackedFiles: trackedFiles(),
      paths: ["scripts", "scripts/orient.ts"],
      diffRef: null,
    });

    expect(files).toEqual([...new Set(files)].toSorted((a, b) => a.localeCompare(b)));
    expect(files).toContain("scripts/orient.ts");
  });

  test("rejects paths outside the repository before reading them", () => {
    expect(() =>
      resolveFiles({
        repoRoot: REPO_ROOT,
        trackedFiles: trackedFiles(),
        paths: ["../private-file"],
        diffRef: null,
      }),
    ).toThrow("Path is outside the repository");
  });

  test("surfaces an invalid diff reference", () => {
    expect(() =>
      resolveFiles({
        repoRoot: REPO_ROOT,
        trackedFiles: trackedFiles(),
        paths: [],
        diffRef: "invalid-orientation-ref",
      }),
    ).toThrow("git diff");
  });

  test("rejects a tracked source symlink whose target is outside the repository", () => {
    const repoRoot = mkdtempSync(path.join(tmpdir(), "folio-orient-repo-"));
    const outsideRoot = mkdtempSync(path.join(tmpdir(), "folio-orient-outside-"));
    try {
      const outsideFile = path.join(outsideRoot, "outside.ts");
      writeFileSync(outsideFile, "export const privateValue = true;\n");
      symlinkSync(outsideFile, path.join(repoRoot, "tracked.ts"));

      expect(() => buildImportGraph(repoRoot, ["tracked.ts"])).toThrow(
        "Path resolves outside the repository",
      );
    } finally {
      rmSync(repoRoot, { recursive: true });
      rmSync(outsideRoot, { recursive: true });
    }
  });

  test("report ordering and aggregate checks are deterministic", () => {
    const files = [
      "parity/compare.ts",
      "packages/core/src/layout-engine/paginator.ts",
      "packages/core/src/docx/ensureParaIds.ts",
    ];
    const tracked = trackedFiles();
    expect(buildReport(REPO_ROOT, tracked, files)).toEqual(buildReport(REPO_ROOT, tracked, files));
  });
});
