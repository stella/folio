#!/usr/bin/env bun
/**
 * Multi-engine DOCX rendering comparison: CLI entry point.
 *
 * Usage:
 *   bun parity/cli.ts                          # full default corpus, HTML report
 *   bun parity/cli.ts path/to/file.docx        # one document
 *   bun parity/cli.ts --help                   # print usage
 *   bun parity/cli.ts some/dir --json          # machine-readable output
 *   bun parity/cli.ts path/to/file.docx --output parity.json
 *   bun parity/cli.ts --reference libreoffice  # open-source reference (default)
 *   bun parity/cli.ts --reference word         # optional proprietary reference
 *   bun parity/cli.ts --refresh-reference      # ignore cached reference exports
 *   bun parity/cli.ts --headed                 # show the folio browser window
 *   bun parity/cli.ts --no-report              # skip writing the HTML report
 *   bun parity/cli.ts path/to/file.docx --max-pages 20
 *
 * Every renderer is an explicit comparison reference, never “ground truth.”
 * LibreOffice is the default. All references require `mutool`
 * (`brew install mupdf-tools`) for normalized PDF geometry extraction.
 */
import { mkdir, stat } from "node:fs/promises";
import path from "node:path";

import { DEFAULT_CORPUS_DIRS } from "./config";
import {
  attributeDivergences,
  clusterCorpus,
  detectFontEnvironment,
  extractDocFeatures,
} from "./features";
import type { ParagraphFeatures } from "./features";
import { createFolioExtractor } from "./folioExtract";
import type { FolioExtractor } from "./folioExtract";
import { compareGeoms } from "./compare";
import { getReferenceRenderer, isReferenceRendererId } from "./referenceRenderer";
import type { ReferenceRenderer } from "./referenceRenderer";
import { writeHtmlReport } from "./report";
import type { DocAssets } from "./report";
import type {
  CorpusReport,
  Divergence,
  DivergenceKind,
  DocGeom,
  FeatureAttributedResult,
  ReferenceRendererId,
} from "./types";

const EXIT_OK = 0;
const EXIT_DIVERGENT = 1;
const EXIT_INFRA_FAILURE = 2;

const TOP_CLUSTER_LIMIT = 10;
const PERFECT_SCORE = 1;
const DEFAULT_REFERENCE_RENDERER: ReferenceRendererId = "libreoffice";

/** Word lock files ("~$name.docx") and dotfiles are never real documents. */
const DOCX_GLOB = "**/*.docx";
const isRealDocxName = (name: string): boolean => !name.startsWith("~$") && !name.startsWith(".");

type CliFlags = {
  help: boolean;
  json: boolean;
  refreshReference: boolean;
  referenceId: ReferenceRendererId;
  headed: boolean;
  reuseServer: boolean;
  noReport: boolean;
  outputPath?: string;
  maxPages?: number;
  paths: string[];
};

export const parseArgs = (argv: string[]): CliFlags => {
  const flags: CliFlags = {
    help: false,
    json: false,
    refreshReference: false,
    referenceId: DEFAULT_REFERENCE_RENDERER,
    headed: false,
    reuseServer: false,
    noReport: false,
    paths: [],
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === undefined) continue;
    if (arg === "--help" || arg === "-h") {
      flags.help = true;
    } else if (arg === "--json") {
      flags.json = true;
    } else if (arg === "--refresh-reference" || arg === "--refresh-truth") {
      flags.refreshReference = true;
    } else if (arg === "--reference") {
      const value = argv[i + 1];
      if (value === undefined || value.startsWith("--")) {
        throw new Error("--reference requires libreoffice or word");
      }
      if (!isReferenceRendererId(value)) {
        throw new Error(`Unknown reference renderer: ${value}`);
      }
      flags.referenceId = value;
      i += 1;
    } else if (arg === "--headed") {
      flags.headed = true;
    } else if (arg === "--reuse-server") {
      flags.reuseServer = true;
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

const HELP_TEXT = `DOCX rendering interoperability comparison

Usage:
  bun parity/cli.ts [doc-or-dir ...] [options]

Options:
  -h, --help           Print this help and exit.
  --json               Print the machine-readable report to stdout.
  --output <path>      Write the JSON report to a file.
  --max-pages <n>      Compare only the first n pages.
  --reference <id>     Reference renderer: libreoffice (default) or word.
  --refresh-reference  Re-render the reference instead of using cache.
  --headed             Show the Folio browser window.
  --reuse-server       Reuse a healthy current-worktree playground server.
  --no-report          Skip writing the HTML report.

Examples:
  bun parity/cli.ts path/to/file.docx --reference libreoffice --max-pages 20
  bun parity/cli.ts path/to/file.docx --reference word
  bun parity/cli.ts some/dir --json --output /tmp/parity.json
`;

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

/** Per-doc pipeline failure: any throw during reference/folio/compare/features. */
type DocFailure = { file: string; error: Error };

/** Result of running the full per-doc pipeline over the corpus. */
type PipelineOutcome = {
  results: FeatureAttributedResult[];
  paragraphsByDoc: ParagraphFeatures[][];
  assets: Map<string, DocAssets>;
  failures: DocFailure[];
};

/**
 * Runs reference-render -> folio-extract -> compare -> feature-attribution
 * for every document sequentially. Some external renderers are single-instance
 * applications and the folio extractor shares one browser page, so documents
 * cannot be processed concurrently.
 */
const runPipeline = async (
  docs: string[],
  flags: CliFlags,
  referenceRenderer: ReferenceRenderer,
): Promise<PipelineOutcome> => {
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
      process.stderr.write(`${label}: ${referenceRenderer.displayName}… `);

      try {
        // oxlint-disable-next-line no-await-in-loop -- references are rendered sequentially for deterministic app automation
        const referenceGeom = limitGeomPages(
          await referenceRenderer.getGeometry(doc, { refresh: flags.refreshReference }),
          flags.maxPages,
        );

        if (!extractor) {
          // oxlint-disable-next-line no-await-in-loop -- created lazily on first use, before any folio extraction
          extractor = await createFolioExtractor({
            headless: !flags.headed,
            reuseServer: flags.reuseServer,
          });
        }

        process.stderr.write("folio… ");
        // oxlint-disable-next-line no-await-in-loop -- the extractor shares one browser page across docs
        const pageLimit = flags.maxPages === undefined ? {} : { maxPages: flags.maxPages };
        const folio = await extractor.extract(doc, pageLimit);

        const result = compareGeoms(referenceGeom, folio.geom);

        // oxlint-disable-next-line no-await-in-loop -- sequential per-doc pipeline
        const docFeatures = await extractDocFeatures(doc);
        const fontEnvironment = detectFontEnvironment(doc, referenceGeom, folio.geom);
        if (fontEnvironment.tags.length > 0) {
          docFeatures.docFeatures.push(...fontEnvironment.tags);
        }
        if (fontEnvironment.status === "shared-substitution") {
          process.stderr.write(
            `(shared substituted font: ${fontEnvironment.tags.join(", ")}; ${fontEnvironment.comparedLines} lines) `,
          );
        } else if (fontEnvironment.status === "mismatch") {
          if (fontEnvironment.tags.includes("font-renderer-metric-mismatch")) {
            process.stderr.write(
              `(${referenceRenderer.displayName}/Folio font metric mismatch despite ${fontEnvironment.matchingLines}/${fontEnvironment.comparedLines} matching family names) `,
            );
          } else {
            process.stderr.write(
              `(${referenceRenderer.displayName}/Folio font mismatch: ${fontEnvironment.matchingLines}/${fontEnvironment.comparedLines} lines match) `,
            );
          }
        } else if (fontEnvironment.status === "unverified") {
          process.stderr.write(`(${referenceRenderer.displayName}/Folio font parity unverified) `);
        }
        const attributed = attributeDivergences(result, docFeatures);

        // oxlint-disable-next-line no-await-in-loop -- sequential per-doc pipeline
        const referencePagePngs = limitPaths(
          await referenceRenderer.getPagePngs(doc, pageLimit),
          flags.maxPages,
        );

        results.push(attributed);
        paragraphsByDoc.push(docFeatures.paragraphs);
        assets.set(doc, {
          referencePagePngs,
          folioPagePngs: folio.screenshotPaths,
          referenceGeom,
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
  const version = report.reference.version ?? "unknown version";
  console.log(
    `\nDOCX interoperability report — ${docCount} document${docCount === 1 ? "" : "s"}, Folio vs ${report.reference.displayName} ${version}\n`,
  );

  for (const result of report.results) {
    const pct = `${(result.score * 100).toFixed(1)}%`;
    const pages = `pages ${result.referencePages}/${result.folioPages}`;
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

  if (flags.help) {
    console.log(HELP_TEXT);
    return EXIT_OK;
  }

  const referenceRenderer = getReferenceRenderer(flags.referenceId);
  if (!(await referenceRenderer.isAvailable())) {
    console.error(
      `${referenceRenderer.displayName} comparison requires the renderer and \`mutool\`.\n` +
        `  - ${referenceRenderer.installHint}\n` +
        "  - Install mutool: brew install mupdf-tools\n" +
        "This is a local, on-demand interoperability comparison; no renderer is treated as specification ground truth.",
    );
    return EXIT_INFRA_FAILURE;
  }

  const docs = await resolveCorpus(flags.paths);
  if (docs.length === 0) {
    console.error("No .docx files found in the given paths (or the default corpus dirs).");
    return EXIT_OK;
  }

  const { results, paragraphsByDoc, assets, failures } = await runPipeline(
    docs,
    flags,
    referenceRenderer,
  );
  const clusters = clusterCorpus(results, paragraphsByDoc);
  const referenceVersion = (await referenceRenderer.getVersion()) ?? undefined;

  const report: CorpusReport = {
    generatedAt: new Date().toISOString(),
    reference: {
      id: referenceRenderer.id,
      displayName: referenceRenderer.displayName,
      ...(referenceVersion !== undefined ? { version: referenceVersion } : {}),
    },
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
    console.error(`DOCX interoperability comparison failed: ${err.name}: ${err.message}`);
    process.exitCode = EXIT_INFRA_FAILURE;
  }
}
