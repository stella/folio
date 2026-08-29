import { describe, expect, test } from "bun:test";

import { isFullyClean, parseArgs } from "../cli";
import type { CorpusReport, FeatureAttributedResult } from "../types";

describe("parity CLI args", () => {
  test("parses --help and -h without treating them as input documents", () => {
    for (const arg of ["--help", "-h"]) {
      const flags = parseArgs([arg]);
      expect(flags.help).toBe(true);
      expect(flags.paths).toEqual([]);
    }
  });

  test("parses an explicit JSON output path without treating it as an input document", () => {
    const flags = parseArgs(["fixture.docx", "--max-pages", "3", "--output", "/tmp/out.json"]);
    expect(flags.paths).toEqual(["fixture.docx"]);
    expect(flags.maxPages).toBe(3);
    expect(flags.outputPath).toBe("/tmp/out.json");
  });

  test("parses the explicit playground server reuse opt-in", () => {
    const flags = parseArgs(["fixture.docx", "--reuse-server"]);
    expect(flags.paths).toEqual(["fixture.docx"]);
    expect(flags.reuseServer).toBe(true);
  });

  test("defaults to LibreOffice and accepts Word as an explicit reference", () => {
    expect(parseArgs(["fixture.docx"]).referenceId).toBe("libreoffice");
    expect(parseArgs(["fixture.docx", "--reference", "word"]).referenceId).toBe("word");
  });

  test("validates the reference renderer", () => {
    expect(() => parseArgs(["fixture.docx", "--reference"])).toThrow(
      "--reference requires libreoffice or word",
    );
    expect(() => parseArgs(["fixture.docx", "--reference", "pages"])).toThrow(
      "Unknown reference renderer: pages",
    );
  });

  test("keeps the old refresh flag as an alias", () => {
    expect(parseArgs(["--refresh-reference"]).refreshReference).toBe(true);
    expect(parseArgs(["--refresh-truth"]).refreshReference).toBe(true);
  });

  test("requires a path after --output", () => {
    expect(() => parseArgs(["fixture.docx", "--output"])).toThrow("--output requires a file path");
    expect(() => parseArgs(["fixture.docx", "--output", "--json"])).toThrow(
      "--output requires a file path",
    );
  });

  test("rejects fractional and partially numeric page limits", () => {
    for (const value of ["1.5", "2pages"]) {
      expect(() => parseArgs(["fixture.docx", "--max-pages", value])).toThrow(
        "--max-pages requires a positive integer",
      );
    }
  });
});

const makeResult = (
  rasterComparison: FeatureAttributedResult["rasterComparison"],
): FeatureAttributedResult => ({
  file: "fixture.docx",
  score: 1,
  referencePages: 1,
  folioPages: 1,
  totalReferenceLines: 0,
  matchedLines: 0,
  medianYOffsetPt: 0,
  divergences: [],
  attributed: [],
  docFeatures: [],
  ...(rasterComparison === undefined ? {} : { rasterComparison }),
});

const makeReport = (result: FeatureAttributedResult): CorpusReport => ({
  generatedAt: "2026-01-01T00:00:00.000Z",
  reference: { id: "libreoffice", displayName: "LibreOffice" },
  results: [result],
  clusters: [],
});

describe("parity clean status", () => {
  test("accepts absent raster diagnostics", () => {
    expect(isFullyClean(makeReport(makeResult(undefined)), [])).toBe(true);
  });

  test("accepts a perfect raster comparison", () => {
    expect(
      isFullyClean(
        makeReport(
          makeResult({ status: "compared", score: 1, diffPixels: 0, totalPixels: 10, pages: [] }),
        ),
        [],
      ),
    ).toBe(true);
  });

  test("rejects empty and imperfect raster diagnostics", () => {
    expect(isFullyClean(makeReport(makeResult({ status: "empty", pages: [] })), [])).toBe(false);
    expect(
      isFullyClean(
        makeReport(
          makeResult({ status: "compared", score: 0.9, diffPixels: 1, totalPixels: 10, pages: [] }),
        ),
        [],
      ),
    ).toBe(false);
  });
});
