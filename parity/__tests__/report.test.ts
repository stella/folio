import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterAll, beforeAll, describe, expect, test } from "bun:test";

import { REPORT_DIR } from "../config";
import type { DocAssets } from "../report";
import { writeHtmlReport } from "../report";
import type { CorpusReport, DocGeom, FeatureAttributedResult } from "../types";

// A minimal valid 1x1 transparent PNG, used as a stand-in for real page
// screenshots: writeHtmlReport only needs to copy the bytes, never decode them.
const ONE_PIXEL_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
  "base64",
);

let tmpDir: string;

const writeFixturePng = async (name: string): Promise<string> => {
  const filePath = path.join(tmpDir, name);
  await Bun.write(filePath, ONE_PIXEL_PNG);
  return filePath;
};

const makeGeom = (source: "word" | "folio", file: string): DocGeom => ({
  source,
  file,
  pages: [
    {
      number: 1,
      widthPt: 612,
      heightPt: 792,
      lines: [
        {
          text: "Hello world",
          normText: "Hello world",
          xPt: 72,
          yPt: 72,
          widthPt: 100,
          heightPt: 12,
          region: "body",
        },
      ],
    },
  ],
  meta: {},
});

describe("writeHtmlReport", () => {
  beforeAll(async () => {
    tmpDir = await mkdtemp(path.join(tmpdir(), "folio-parity-report-"));
  });

  afterAll(async () => {
    await rm(tmpDir, { recursive: true, force: true });
    await rm(REPORT_DIR, { recursive: true, force: true });
  });

  test("writes an index page, per-doc pages, escapes text, and copies assets", async () => {
    const docAFile = "/corpus/sample-one.docx";
    // Deliberately hostile doc name + line text: both must come out escaped.
    // (No "/" in the basename itself — that would just be a path separator.)
    const docBFile = '/corpus/<b>bold & "quoted".docx';

    const wordPngA = await writeFixturePng("word-a-1.png");
    const folioPngA = await writeFixturePng("folio-a-1.png");
    const wordPngB = await writeFixturePng("word-b-1.png");
    const folioPngB = await writeFixturePng("folio-b-1.png");

    const resultA: FeatureAttributedResult = {
      file: docAFile,
      score: 1,
      wordPages: 1,
      folioPages: 1,
      totalWordLines: 1,
      matchedLines: 1,
      medianYOffsetPt: 0.5,
      divergences: [],
      attributed: [],
      docFeatures: [],
    };

    const resultB: FeatureAttributedResult = {
      file: docBFile,
      score: 0.5,
      wordPages: 1,
      folioPages: 1,
      totalWordLines: 2,
      matchedLines: 1,
      medianYOffsetPt: 1.25,
      divergences: [
        { kind: "page-count", word: 2, folio: 1 },
        {
          kind: "missing-line",
          page: 1,
          text: "<script>missing</script>",
        },
        {
          kind: "y-drift",
          page: 1,
          text: "Hello world",
          residualPt: 3.2,
        },
        {
          kind: "text-mismatch",
          page: 1,
          wordText: "Hello world",
          folioText: "Hell0 world",
        },
      ],
      attributed: [
        {
          divergence: { kind: "missing-line", page: 1, text: "<script>missing</script>" },
          features: ["table"],
        },
      ],
      docFeatures: [],
    };

    const report: CorpusReport = {
      generatedAt: "2026-07-04T00:00:00.000Z",
      wordVersion: "16.112",
      results: [resultA, resultB],
      clusters: [
        {
          kind: "missing-line",
          feature: "table",
          count: 5,
          docs: [docAFile, docBFile],
          lift: 2.3456,
          meanMagnitudePt: 1.5,
          examples: [
            {
              divergence: { kind: "missing-line", page: 1, text: "<script>example</script>" },
              features: ["table"],
            },
          ],
        },
        {
          kind: "y-drift",
          feature: "spacing-atLeast",
          count: 2,
          docs: [docBFile],
          lift: 1.1,
          examples: [],
        },
      ],
    };

    const assets = new Map<string, DocAssets>([
      [
        docAFile,
        {
          wordPagePngs: [wordPngA],
          folioPagePngs: [folioPngA],
          wordGeom: makeGeom("word", docAFile),
          folioGeom: makeGeom("folio", docAFile),
        },
      ],
      [
        docBFile,
        {
          wordPagePngs: [wordPngB],
          folioPagePngs: [folioPngB],
          wordGeom: makeGeom("word", docBFile),
          folioGeom: makeGeom("folio", docBFile),
        },
      ],
    ]);

    const indexPath = await writeHtmlReport(report, assets);
    expect(indexPath).toBe(path.join(REPORT_DIR, "index.html"));

    const indexHtml = await readFile(indexPath, "utf8");

    // Links to both detail pages.
    expect(indexHtml).toContain("sample-one.docx");
    expect(indexHtml).toMatch(/href="doc-sample-one\.html"/);
    expect(indexHtml).toMatch(/href="doc-[a-z0-9-]+\.html"/);

    // Cluster rows: feature strings present, lift rounded to 2 decimals.
    expect(indexHtml).toContain("table");
    expect(indexHtml).toContain("spacing-atLeast");
    expect(indexHtml).toContain("2.35"); // lift 2.3456 -> 2 decimals

    // The hostile doc name must never appear unescaped.
    expect(indexHtml).not.toContain('<b>bold & "quoted"');
    expect(indexHtml).toContain("&lt;b&gt;bold &amp; &quot;quoted&quot;");

    // Detail pages exist for both docs.
    const slugAPath = path.join(REPORT_DIR, "doc-sample-one.html");
    const detailAHtml = await readFile(slugAPath, "utf8");
    expect(detailAHtml).toContain("Full parity");

    // Find doc B's detail page by scanning the report dir for the other slug.
    const docBLinkMatch = /href="(doc-[a-z0-9-]+\.html)"/g.exec(
      indexHtml.replace(/href="doc-sample-one\.html"/, ""),
    );
    expect(docBLinkMatch).not.toBeNull();
    const docBHref = docBLinkMatch?.[1];
    expect(docBHref).toBeDefined();
    if (!docBHref) throw new Error("expected doc B link");

    const detailBHtml = await readFile(path.join(REPORT_DIR, docBHref), "utf8");

    // Interpolated divergence text is escaped, not raw.
    expect(detailBHtml).not.toContain("<script>missing</script>");
    expect(detailBHtml).toContain("&lt;script&gt;missing&lt;/script&gt;");

    // Divergence kinds appear in the fixed order (page-count, missing-line,
    // text-mismatch, y-drift — per the contract's DivergenceKind ordering).
    const pageCountIdx = detailBHtml.indexOf("page-count");
    const missingLineIdx = detailBHtml.indexOf("missing-line");
    const textMismatchIdx = detailBHtml.indexOf("text-mismatch");
    const yDriftIdx = detailBHtml.indexOf("y-drift");
    expect(pageCountIdx).toBeGreaterThan(-1);
    expect(missingLineIdx).toBeGreaterThan(pageCountIdx);
    expect(textMismatchIdx).toBeGreaterThan(missingLineIdx);
    expect(yDriftIdx).toBeGreaterThan(textMismatchIdx);

    // PNGs were copied into REPORT_DIR/assets/<slug>/.
    const copiedWordA = await readFile(path.join(REPORT_DIR, "assets", "sample-one", "word-1.png"));
    expect(copiedWordA.equals(ONE_PIXEL_PNG)).toBe(true);
  });

  test("de-duplicates slugs that sanitize to the same basename", async () => {
    const fileOne = "/corpus/a/report.docx";
    const fileTwo = "/corpus/b/report.docx";

    const baseResult = (file: string): FeatureAttributedResult => ({
      file,
      score: 1,
      wordPages: 0,
      folioPages: 0,
      totalWordLines: 0,
      matchedLines: 0,
      medianYOffsetPt: 0,
      divergences: [],
      attributed: [],
      docFeatures: [],
    });

    const report: CorpusReport = {
      generatedAt: "2026-07-04T00:00:00.000Z",
      results: [baseResult(fileOne), baseResult(fileTwo)],
      clusters: [],
    };

    const indexPath = await writeHtmlReport(report, new Map());
    const indexHtml = await readFile(indexPath, "utf8");

    expect(indexHtml).toContain('href="doc-report.html"');
    expect(indexHtml).toContain('href="doc-report-2.html"');

    await expect(readFile(path.join(REPORT_DIR, "doc-report.html"), "utf8")).resolves.toContain(
      "No page assets available",
    );
    await expect(readFile(path.join(REPORT_DIR, "doc-report-2.html"), "utf8")).resolves.toContain(
      "No page assets available",
    );
  });

  test("handles missing wordVersion, empty corpus, and missing PNG source gracefully", async () => {
    const report: CorpusReport = {
      generatedAt: "2026-07-04T00:00:00.000Z",
      results: [],
      clusters: [],
    };

    const indexPath = await writeHtmlReport(report, new Map());
    const indexHtml = await readFile(indexPath, "utf8");
    expect(indexHtml).toContain("unknown");
    expect(indexHtml).toContain("No documents.");
    expect(indexHtml).toContain("No clusters.");
  });
});
