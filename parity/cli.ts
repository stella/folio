#!/usr/bin/env bun
/**
 * Word rendering parity engine: CLI entry point.
 *
 * Usage:
 *   bun parity/cli.ts                          # full default corpus, HTML report
 *   bun parity/cli.ts path/to/file.docx        # one document
 *   bun parity/cli.ts some/dir --json          # machine-readable output
 *   bun parity/cli.ts path/to/file.docx --output parity.json
 *   bun parity/cli.ts --refresh-truth          # ignore cached Word exports
 *   bun parity/cli.ts --headed                 # show the folio browser window
 *   bun parity/cli.ts --no-report              # skip writing the HTML report
 *   bun parity/cli.ts path/to/file.docx --max-pages 20
 *
 * Ground truth requires Microsoft Word for Mac plus `mutool`
 * (`brew install mupdf-tools`); the CLI exits 2 with a clear message when
 * either is missing, mirroring the differential-test skip ergonomics.
 */
import { mkdir, stat } from "node:fs/promises";
import path from "node:path";

import { DEFAULT_CORPUS_DIRS } from "./config";
import {
  attributeDivergences,
  clusterCorpus,
  detectFontSubstitution,
  extractDocFeatures,
} from "./features";
import type { ParagraphFeatures } from "./features";
import { createFolioExtractor } from "./folioExtract";
import type { FolioExtractor } from "./folioExtract";
import { compareGeoms } from "./compare";
import { writeHtmlReport } from "./report";
import type { DocAssets } from "./report";
import { getWordPagePngs, getWordTruth, getWordVersion, isWordAvailable } from "./wordTruth";
import type {
  CorpusReport,
  Divergence,
  DivergenceKind,
  DocGeom,
  FeatureAttributedResult,
} from "./types";

const EXIT_OK = 0;
const EXIT_DIVERGENT = 1;
const EXIT_INFRA_FAILURE = 2;

const TOP_CLUSTER_LIMIT = 10;
const PERFECT_SCORE = 1;

/** Word lock files ("~$name.docx") and dotfiles are never real documents. */
const DOCX_GLOB = "**/*.docx";
const isRealDocxName = (name: string): boolean => !name.startsWith("~$") && !name.startsWith(".");

type CliFlags = {
  json: boolean;
  refreshTruth: boolean;
  headed: boolean;
  noReport: boolean;
  outputPath?: string;
  maxPages?: number;
  paths: string[];
};

export const parseArgs = (argv: string[]): CliFlags => {
  const flags: CliFlags = {
    json: false,
    refreshTruth: false,
    headed: false,
    noReport: false,
    paths: [],
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--json") {
      flags.json = true;
    } else if (arg === "--refresh-truth") {
      flags.refreshTruth = true;
    } else if (arg === "--headed") {
      flags.headed = true;
    } else if (arg === "--no-report") {
      flags.noReport = true;
    } else if (arg === "--output") {
      const value = argv[i + 1];
      if (!value || value.startsWith("--")) {
        throw new Error("--output requires a file path");
      }
      flags.outputPath = value;
      i += 1;
    } else if (arg === "--max-pages") {
      const value = argv[i + 1];
      const maxPages = value === undefined ? Number.NaN : Number.parseInt(value, 10);
      if (!Number.isInteger(maxPages) || maxPages <= 0) {
        throw new Error("--max-pages requires a positive integer");
      }
      flags.maxPages = maxPages;
      i += 1;
    } else {
      flags.paths.push(arg);
    }
  }
  return flags;
};

const limitGeomPages = (geom: DocGeom, maxPages: number | undefined): DocGeom => {
  if (maxPages === undefined) {
    return geom;
  }
  return { ...geom, pages: geom.pages.slice(0, maxPages) };
};

const limitPaths = (paths: string[], maxPages: number | undefined): string[] =>
  maxPages === undefined ? paths : paths.slice(0, maxPages);

const isDirectory = async (candidate: string): Promise<boolean> => {
  try {
    return (await stat(candidate)).isDirectory();
  } catch {
    return false;
  }
};

const scanDirForDocx = async (dir: string): Promise<string[]> => {
  const glob = new Bun.Glob(DOCX_GLOB);
  const found: string[] = [];
  for await (const rel of glob.scan({ cwd: dir, dot: false })) {
    if (isRealDocxName(path.basename(rel))) {
      found.push(path.resolve(dir, rel));
    }
  }
  return found;
};

/** Resolves CLI path args (or the default corpus dirs) into an absolute,
 * deduplicated, sorted list of .docx files. A directory is scanned
 * recursively for *.docx; a file is taken as-is. */
const resolveCorpus = async (inputPaths: string[]): Promise<string[]> => {
  const roots = inputPaths.length > 0 ? inputPaths : DEFAULT_CORPUS_DIRS;
  const files = new Set<string>();

  for (const root of roots) {
    const absRoot = path.resolve(root);
    // oxlint-disable-next-line no-await-in-loop -- roots are a small, fixed list; scanning is inherently sequential per root
    if (await isDirectory(absRoot)) {
      // oxlint-disable-next-line no-await-in-loop -- see above
      for (const file of await scanDirForDocx(absRoot)) {
        files.add(file);
      }
    } else {
      files.add(absRoot);
    }
  }

  return Array.from(files).toSorted((a, b) => a.localeCompare(b));
};

/** Per-doc pipeline failure: any throw during word-truth/folio/compare/features. */
type DocFailure = { file: string; error: Error };

/** Result of running the full per-doc pipeline over the corpus. */
type PipelineOutcome = {
  results: FeatureAttributedResult[];
  paragraphsByDoc: ParagraphFeatures[][];
  assets: Map<string, DocAssets>;
  failures: DocFailure[];
};

/**
 * Runs word-truth -> folio-extract -> compare -> feature-attribution for
 * every doc in `docs`, sequentially: Word is a single-instance app and the
 * folio extractor shares one browser page, so docs cannot be processed
 * concurrently. A single `FolioExtractor` is created lazily before the first
 * folio extraction and closed in `finally` regardless of how the loop ends.
 */
const runPipeline = async (docs: string[], flags: CliFlags): Promise<PipelineOutcome> => {
  const results: FeatureAttributedResult[] = [];
  const paragraphsByDoc: ParagraphFeatures[][] = [];
  const assets = new Map<string, DocAssets>();
  const failures: DocFailure[] = [];

  let extractor: FolioExtractor | undefined;

  try {
    for (let i = 0; i < docs.length; i++) {
      const doc = docs[i];
      if (!doc) continue;

      const label = `[${i + 1}/${docs.length}] ${path.basename(doc)}`;
      process.stderr.write(`${label}: word truth… `);

      try {
        // oxlint-disable-next-line no-await-in-loop -- Word is a single-instance app; docs are exported one at a time
        const wordGeom = limitGeomPages(
          await getWordTruth(doc, { refresh: flags.refreshTruth }),
          flags.maxPages,
        );

        if (!extractor) {
          // oxlint-disable-next-line no-await-in-loop -- created lazily on first use, before any folio extraction
          extractor = await createFolioExtractor({ headless: !flags.headed });
        }

        process.stderr.write("folio… ");
        // oxlint-disable-next-line no-await-in-loop -- the extractor shares one browser page across docs
        const folio = await extractor.extract(doc, { maxPages: flags.maxPages });

        const result = compareGeoms(wordGeom, folio.geom);

        // oxlint-disable-next-line no-await-in-loop -- sequential per-doc pipeline
        const docFeatures = await extractDocFeatures(doc);
        // Ground truth rendered with substituted fonts (doc requests a font
        // this machine's Word lacks) is a first-class doc-level condition:
        // its divergences may be font availability, not folio layout bugs.
        const substitutionTags = detectFontSubstitution(doc, wordGeom);
        if (substitutionTags.length > 0) {
          docFeatures.docFeatures.push(...substitutionTags);
          process.stderr.write(`(ground truth font-substituted: ${substitutionTags.join(", ")}) `);
        }
        const attributed = attributeDivergences(result, docFeatures);

        // oxlint-disable-next-line no-await-in-loop -- sequential per-doc pipeline
        const wordPagePngs = limitPaths(
          await getWordPagePngs(doc, { maxPages: flags.maxPages }),
          flags.maxPages,
        );

        results.push(attributed);
        paragraphsByDoc.push(docFeatures.paragraphs);
        assets.set(doc, {
          wordPagePngs,
          folioPagePngs: folio.screenshotPaths,
          wordGeom,
          folioGeom: folio.geom,
        });

        process.stderr.write(`score ${result.score.toFixed(2)}\n`);
      } catch (error) {
        const err = error instanceof Error ? error : new Error(String(error));
        process.stderr.write(`\n${label}: FAILED — ${err.name}: ${err.message}\n`);
        failures.push({ file: doc, error: err });
      }
    }
  } finally {
    if (extractor) {
      await extractor.close();
    }
  }

  return { results, paragraphsByDoc, assets, failures };
};

/** Report/CLI-summary divergence ordering, per the fixed contract. */
const DIVERGENCE_KIND_ORDER: DivergenceKind[] = [
  "page-count",
  "pagination",
  "missing-line",
  "extra-line",
  "line-break",
  "text-mismatch",
  "y-drift",
  "x-drift",
  "width-drift",
];

const compactDivergenceCounts = (divergences: Divergence[]): string => {
  const counts = new Map<DivergenceKind, number>();
  for (const divergence of divergences) {
    counts.set(divergence.kind, (counts.get(divergence.kind) ?? 0) + 1);
  }
  return DIVERGENCE_KIND_ORDER.filter((kind) => (counts.get(kind) ?? 0) > 0)
    .map((kind) => `${counts.get(kind)} ${kind}`)
    .join(", ");
};

const printHumanSummary = (report: CorpusReport, failures: DocFailure[]): void => {
  const docCount = report.results.length;
  console.log(
    `\nWord rendering parity report — ${docCount} document${docCount === 1 ? "" : "s"}, Word ${report.wordVersion ?? "unknown"}\n`,
  );

  for (const result of report.results) {
    const pct = `${(result.score * 100).toFixed(1)}%`;
    const pages = `pages ${result.wordPages}/${result.folioPages}`;
    const counts = compactDivergenceCounts(result.divergences) || "no divergences";
    console.log(
      `  ${path.basename(result.file).padEnd(40)} ${pct.padStart(6)}  ${pages}  ${counts}`,
    );
  }

  console.log(`\nTop ${TOP_CLUSTER_LIMIT} clusters:`);
  const topClusters = report.clusters.slice(0, TOP_CLUSTER_LIMIT);
  if (topClusters.length === 0) {
    console.log("  (none)");
  } else {
    for (const cluster of topClusters) {
      console.log(
        `  ${cluster.kind.padEnd(14)} ${cluster.feature.padEnd(24)} count=${cluster.count} lift=${cluster.lift.toFixed(2)}`,
      );
    }
  }

  if (failures.length > 0) {
    console.log("\nFailed docs:");
    for (const failure of failures) {
      console.log(
        `  ${path.basename(failure.file)}: ${failure.error.name}: ${failure.error.message}`,
      );
    }
  }
};

const isFullyClean = (report: CorpusReport, failures: DocFailure[]): boolean =>
  failures.length === 0 &&
  report.results.every(
    (result) => result.score === PERFECT_SCORE && result.divergences.length === 0,
  );

const writeJsonReport = async (report: CorpusReport, outputPath: string): Promise<string> => {
  const resolved = path.resolve(outputPath);
  await mkdir(path.dirname(resolved), { recursive: true });
  await Bun.write(resolved, `${JSON.stringify(report, null, 2)}\n`);
  return resolved;
};

const main = async (argv: string[]): Promise<number> => {
  const flags = parseArgs(argv);

  if (!(await isWordAvailable())) {
    console.error(
      "Word rendering parity engine requires Microsoft Word for Mac and `mutool`.\n" +
        "  - Install Word: https://www.microsoft.com/microsoft-365/word\n" +
        "  - Install mutool: brew install mupdf-tools\n" +
        "This tool is a local, on-demand comparison; it is not a CI gate.",
    );
    return EXIT_INFRA_FAILURE;
  }

  const docs = await resolveCorpus(flags.paths);
  if (docs.length === 0) {
    console.error("No .docx files found in the given paths (or the default corpus dirs).");
    return EXIT_OK;
  }

  const { results, paragraphsByDoc, assets, failures } = await runPipeline(docs, flags);
  const clusters = clusterCorpus(results, paragraphsByDoc);
  const wordVersion = (await getWordVersion()) ?? undefined;

  const report: CorpusReport = {
    generatedAt: new Date().toISOString(),
    ...(wordVersion !== undefined ? { wordVersion } : {}),
    results,
    clusters,
  };

  if (flags.outputPath) {
    const outputPath = await writeJsonReport(report, flags.outputPath);
    console.error(`JSON: ${outputPath}`);
  }

  if (flags.json && !flags.outputPath) {
    console.log(JSON.stringify(report, null, 2));
  } else if (!flags.json) {
    printHumanSummary(report, failures);
  }

  if (!flags.noReport) {
    const indexPath = await writeHtmlReport(report, assets);
    if (flags.json) {
      console.error(`Report: ${indexPath}`);
    } else {
      console.log(`\nReport: ${indexPath}`);
    }
  }

  return isFullyClean(report, failures) ? EXIT_OK : EXIT_DIVERGENT;
};

if (import.meta.main) {
  try {
    process.exitCode = await main(process.argv.slice(2));
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    console.error(`Word rendering parity engine failed: ${err.name}: ${err.message}`);
    process.exitCode = EXIT_INFRA_FAILURE;
  }
}
