import { describe, expect, test } from "bun:test";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { LayoutInteractionCase } from "../fixtures/layout-interaction-matrix";
import { LayoutMatrixError, type ReferenceLayoutMatrixReport } from "../layoutMatrix";
import { parseLayoutMatrixCliArgs, writeLayoutMatrixArtifacts } from "../layoutMatrixCli";

describe("reference layout matrix CLI", () => {
  test("defaults to a cached headless run", () => {
    expect(parseLayoutMatrixCliArgs([])).toMatchObject({
      type: "run",
      refreshReference: false,
      referenceId: "libreoffice",
      headed: false,
      reuseServer: false,
    });
  });

  test("parses run controls", () => {
    expect(
      parseLayoutMatrixCliArgs([
        "--reference",
        "word",
        "--refresh-reference",
        "--headed",
        "--reuse-server",
        "--output",
        "/tmp/matrix.json",
      ]),
    ).toEqual({
      type: "run",
      refreshReference: true,
      referenceId: "word",
      headed: true,
      reuseServer: true,
      outputPath: "/tmp/matrix.json",
    });
  });

  test("supports help and rejects incomplete flags", () => {
    expect(parseLayoutMatrixCliArgs(["--help"])).toEqual({ type: "help" });
    expect(() => parseLayoutMatrixCliArgs(["--output"])).toThrow(LayoutMatrixError);
    expect(() => parseLayoutMatrixCliArgs(["--reference", "pages"])).toThrow(
      "Unknown reference renderer: pages",
    );
  });

  test("writes JSON after the visual report recreates its directory", async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "layout-matrix-cli-test-"));
    const outputPath = path.join(outputDirectory, "layout-matrix.json");
    try {
      const scenario = {
        id: "mx-synthetic",
        section: "single",
        anchorFrame: "inline",
        wrap: "inline",
        flow: "normal",
        table: "none",
        typography: "latin",
      } satisfies LayoutInteractionCase;
      const report = {
        schema: "folio.reference-layout-matrix",
        version: 1,
        generatedAt: "2026-09-09T00:00:00.000Z",
        reference: { id: "word", displayName: "Word" },
        summary: { total: 1, pass: 1, advisory: 0, fail: 0 },
        cases: [
          {
            scenario,
            status: "pass",
            geometryReliable: true,
            referencePages: 1,
            folioPages: 1,
            requiredDivergences: {},
            diagnosticDivergences: {},
            requiredRasterFailures: 0,
          },
        ],
        requiredInteractionClusters: [],
      } satisfies ReferenceLayoutMatrixReport;
      await writeLayoutMatrixArtifacts({
        report,
        outputPath,
        writeVisualReport: async () => {
          await rm(outputDirectory, { recursive: true, force: true });
          await mkdir(outputDirectory, { recursive: true });
          return path.join(outputDirectory, "index.html");
        },
      });

      expect(await Bun.file(outputPath).exists()).toBeTrue();
    } finally {
      await rm(outputDirectory, { recursive: true, force: true });
    }
  });
});
