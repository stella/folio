#!/usr/bin/env bun
/** Run every synthetic layout interaction case against a selected reference renderer. */
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { runPipeline, type RunPipelineOptions } from "./cli";
import { clusterCorpus } from "./features";
import { buildLayoutInteractionCaseFixture } from "./fixtures/build-layout-corpus";
import { buildLayoutInteractionMatrix } from "./fixtures/layout-interaction-matrix";
import { layoutMatrixError, summarizeLayoutMatrix } from "./layoutMatrix";
import { REPORT_DIR } from "./config";
import { getReferenceRenderer, isReferenceRendererId } from "./referenceRenderer";
import { writeHtmlReport } from "./report";
import type { CorpusReport, ReferenceRendererId } from "./types";

const EXIT_OK = 0;
const EXIT_DIVERGENT = 1;
const EXIT_INFRA_FAILURE = 2;
const DEFAULT_REFERENCE: ReferenceRendererId = "libreoffice";
const TOP_CLUSTER_LIMIT = 10;

export type LayoutMatrixCliCommand =
  | { type: "help" }
  | {
      type: "run";
      refreshReference: boolean;
      referenceId: ReferenceRendererId;
      headed: boolean;
      reuseServer: boolean;
      outputPath: string;
    };

const requireFlagValue = (argv: string[], index: number, flag: string): string => {
  const value = argv[index + 1];
  if (value === undefined || value.startsWith("--")) {
    throw layoutMatrixError(`${flag} requires a path.`);
  }
  return value;
};

export const parseLayoutMatrixCliArgs = (argv: string[]): LayoutMatrixCliCommand => {
  if (argv.includes("--help") || argv.includes("-h")) return { type: "help" };
  let refreshReference = false;
  let referenceId: ReferenceRendererId = DEFAULT_REFERENCE;
  let headed = false;
  let reuseServer = false;
  let outputPath: string | undefined;
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--refresh-reference") {
      refreshReference = true;
    } else if (arg === "--reference") {
      const value = requireFlagValue(argv, index, arg);
      if (!isReferenceRendererId(value)) {
        throw layoutMatrixError(`Unknown reference renderer: ${value}.`);
      }
      referenceId = value;
      index += 1;
    } else if (arg === "--headed") {
      headed = true;
    } else if (arg === "--reuse-server") {
      reuseServer = true;
    } else if (arg === "--output") {
      outputPath = requireFlagValue(argv, index, arg);
      index += 1;
    } else {
      throw layoutMatrixError(`Unknown option: ${String(arg)}`);
    }
  }
  return {
    type: "run",
    refreshReference,
    referenceId,
    headed,
    reuseServer,
    outputPath: outputPath ?? path.join(REPORT_DIR, `layout-matrix-${referenceId}.json`),
  };
};

const HELP_TEXT = `Synthetic reference-renderer layout matrix

Usage:
  bun parity/layoutMatrixCli.ts [--reference <id>] [--refresh-reference] [--headed] [--reuse-server] [--output <report.json>]

Each generated case contains invented values only. Reference exports are cached
by fixture hash; pass --refresh-reference after changing the environment. Page,
page-size, and reliable line-flow failures make the command fail. Raw pixel
differences and geometry drift remain advisory diagnostics.
`;

const writeCaseCorpus = async (directory: string): Promise<string[]> => {
  const paths: string[] = [];
  const scenarios = buildLayoutInteractionMatrix();
  for (const scenario of scenarios) {
    const filePath = path.join(directory, `${scenario.id}.docx`);
    // oxlint-disable-next-line no-await-in-loop -- deterministic fixture generation is bounded to the matrix size
    await Bun.write(filePath, await buildLayoutInteractionCaseFixture(scenario));
    paths.push(filePath);
  }
  return paths;
};

const printSummary = (report: ReturnType<typeof summarizeLayoutMatrix>): void => {
  const { summary } = report;
  console.log(
    `\nLayout matrix: ${summary.total} cases, ${summary.fail} required failure${summary.fail === 1 ? "" : "s"}, ${summary.advisory} advisory, ${summary.pass} exact.`,
  );
  if (report.requiredInteractionClusters.length === 0) return;
  console.log("Required interaction clusters:");
  for (const cluster of report.requiredInteractionClusters.slice(0, TOP_CLUSTER_LIMIT)) {
    console.log(
      `  ${cluster.pair}: ${cluster.failingCases} case${cluster.failingCases === 1 ? "" : "s"}`,
    );
  }
};

type WriteLayoutMatrixArtifactsOptions = {
  report: ReturnType<typeof summarizeLayoutMatrix>;
  outputPath: string;
  writeVisualReport: () => Promise<string>;
};

export const writeLayoutMatrixArtifacts = async ({
  report,
  outputPath,
  writeVisualReport,
}: WriteLayoutMatrixArtifactsOptions): Promise<{ htmlPath: string; outputPath: string }> => {
  const htmlPath = await writeVisualReport();
  const resolvedOutputPath = path.resolve(outputPath);
  await mkdir(path.dirname(resolvedOutputPath), { recursive: true });
  await Bun.write(resolvedOutputPath, `${JSON.stringify(report, null, 2)}\n`);
  return { htmlPath, outputPath: resolvedOutputPath };
};

const run = async (command: Extract<LayoutMatrixCliCommand, { type: "run" }>): Promise<number> => {
  const renderer = getReferenceRenderer(command.referenceId);
  if (!(await renderer.isAvailable())) {
    throw layoutMatrixError(`${renderer.displayName} and mutool are required.`);
  }
  const workDirectory = await mkdtemp(path.join(tmpdir(), "folio-layout-matrix-"));
  try {
    const docs = await writeCaseCorpus(workDirectory);
    const flags: RunPipelineOptions = {
      refreshReference: command.refreshReference,
      referenceId: command.referenceId,
      headed: command.headed,
      reuseServer: command.reuseServer,
    };
    const { results, paragraphsByDoc, assets, failures } = await runPipeline(docs, flags, renderer);
    if (failures.length > 0) {
      const ids = failures.map(({ file }) => path.basename(file, ".docx")).join(", ");
      throw layoutMatrixError(`Matrix infrastructure failed for: ${ids}.`);
    }
    const referenceVersion = (await renderer.getVersion()) ?? undefined;
    const corpusReport: CorpusReport = {
      generatedAt: new Date().toISOString(),
      reference: {
        id: renderer.id,
        displayName: renderer.displayName,
        ...(referenceVersion === undefined ? {} : { version: referenceVersion }),
      },
      results,
      clusters: clusterCorpus(results, paragraphsByDoc),
    };
    const report = summarizeLayoutMatrix(corpusReport, buildLayoutInteractionMatrix());
    const { htmlPath, outputPath } = await writeLayoutMatrixArtifacts({
      report,
      outputPath: command.outputPath,
      writeVisualReport: async () => await writeHtmlReport(corpusReport, assets),
    });
    printSummary(report);
    console.log(`Matrix JSON: ${outputPath}`);
    console.log(`Visual report: ${htmlPath}`);
    return report.summary.fail === 0 ? EXIT_OK : EXIT_DIVERGENT;
  } finally {
    await rm(workDirectory, { recursive: true, force: true });
  }
};

export const runLayoutMatrixCli = async (argv: string[]): Promise<number> => {
  const command = parseLayoutMatrixCliArgs(argv);
  if (command.type === "help") {
    console.log(HELP_TEXT);
    return EXIT_OK;
  }
  return await run(command);
};

if (import.meta.main) {
  try {
    process.exitCode = await runLayoutMatrixCli(process.argv.slice(2));
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    console.error(`Layout matrix failed: ${err.name}: ${err.message}`);
    process.exitCode = EXIT_INFRA_FAILURE;
  }
}
