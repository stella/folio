import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { PNG } from "pngjs";

import { comparePageRasters, RasterComparisonError } from "../rasterCompare";

let tmpDir: string;

const writeSolidPng = async (
  name: string,
  width: number,
  height: number,
  color: readonly [number, number, number, number],
): Promise<string> => {
  const png = new PNG({ width, height });
  for (let offset = 0; offset < png.data.length; offset += 4) {
    png.data[offset] = color[0];
    png.data[offset + 1] = color[1];
    png.data[offset + 2] = color[2];
    png.data[offset + 3] = color[3];
  }
  const filePath = path.join(tmpDir, name);
  await Bun.write(filePath, PNG.sync.write(png));
  return filePath;
};

describe("page raster comparison", () => {
  beforeAll(async () => {
    tmpDir = await mkdtemp(path.join(tmpdir(), "folio-raster-compare-"));
  });

  afterAll(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  test("reports identical pages without pixel differences", async () => {
    const reference = await writeSolidPng("same-reference.png", 2, 2, [255, 255, 255, 255]);
    const folio = await writeSolidPng("same-folio.png", 2, 2, [255, 255, 255, 255]);
    const outputDir = path.join(tmpDir, "same-diffs");

    const result = await comparePageRasters({
      referencePagePngs: [reference],
      folioPagePngs: [folio],
      outputDir,
    });

    expect(result.comparison).toEqual({
      status: "compared",
      score: 1,
      diffPixels: 0,
      totalPixels: 4,
      pages: [
        {
          status: "match",
          page: 1,
          widthPx: 2,
          heightPx: 2,
          diffPixels: 0,
          totalPixels: 4,
          similarity: 1,
        },
      ],
    });
    expect(await readFile(result.diffPagePngs[0] ?? "")).not.toHaveLength(0);
  });

  test("counts painted differences across equal-size pages", async () => {
    const reference = await writeSolidPng("different-reference.png", 2, 2, [255, 255, 255, 255]);
    const folio = await writeSolidPng("different-folio.png", 2, 2, [0, 0, 0, 255]);

    const { comparison } = await comparePageRasters({
      referencePagePngs: [reference],
      folioPagePngs: [folio],
      outputDir: path.join(tmpDir, "different-diffs"),
    });

    expect(comparison.status).toBe("compared");
    if (comparison.status !== "compared") return;
    expect(comparison.score).toBe(0);
    expect(comparison.pages[0]?.status).toBe("difference");
    expect(comparison.diffPixels).toBe(4);
  });

  test("makes page dimensions part of the result", async () => {
    const reference = await writeSolidPng("size-reference.png", 2, 2, [255, 255, 255, 255]);
    const folio = await writeSolidPng("size-folio.png", 1, 2, [255, 255, 255, 255]);

    const { comparison } = await comparePageRasters({
      referencePagePngs: [reference],
      folioPagePngs: [folio],
      outputDir: path.join(tmpDir, "size-diffs"),
    });

    expect(comparison.status).toBe("compared");
    if (comparison.status !== "compared") return;
    expect(comparison.pages[0]).toMatchObject({
      status: "dimension-mismatch",
      referenceWidthPx: 2,
      referenceHeightPx: 2,
      folioWidthPx: 1,
      folioHeightPx: 2,
      similarity: 0.5,
    });
    expect(comparison.score).toBe(0.5);
  });

  test("treats a missing page as fully different", async () => {
    const reference = await writeSolidPng("missing-reference.png", 2, 2, [255, 255, 255, 255]);

    const { comparison } = await comparePageRasters({
      referencePagePngs: [reference],
      folioPagePngs: [],
      outputDir: path.join(tmpDir, "missing-diffs"),
    });

    expect(comparison.status).toBe("compared");
    if (comparison.status !== "compared") return;
    expect(comparison.score).toBe(0);
    expect(comparison.pages[0]?.status).toBe("missing-folio");
  });

  test("rejects an excessive page count before creating output", async () => {
    const pagePaths = Array.from({ length: 251 }, () => "unused.png");
    const outputDir = path.join(tmpDir, "excessive-page-diffs");

    await expect(
      comparePageRasters({
        referencePagePngs: pagePaths,
        folioPagePngs: [],
        outputDir,
      }),
    ).rejects.toBeInstanceOf(RasterComparisonError);
    expect(await Bun.file(outputDir).exists()).toBe(false);
  });
});
