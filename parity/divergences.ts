#!/usr/bin/env bun
/**
 * Static parity-report slicer for agent diagnostics.
 *
 * Reads an existing `parity/cli.ts --json` report and prints a deterministic,
 * filterable list of divergences plus ready-to-run follow-up commands. This
 * avoids re-running the reference renderer and Folio when the next step is
 * just choosing a target.
 */
import path from "node:path";

import { isReferenceRendererId } from "./referenceRenderer";
import { normalizeLineText } from "./textNorm";
import type { CorpusReport, Divergence, DivergenceKind } from "./types";

type Flags = {
  report?: string | undefined;
  doc?: string | undefined;
  page?: number | undefined;
  kind?: DivergenceKind | undefined;
  text?: string | undefined;
  limit: number;
};

const KIND_ORDER: DivergenceKind[] = [
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

const usage = (): string =>
  [
    "Usage:",
    "  bun parity/divergences.ts --report /tmp/run.json --page 3 --limit 20",
    "  bun parity/divergences.ts --report /tmp/run.json --kind text-mismatch --text 'Schedule 1'",
    "",
    "Options:",
    "  --report <path>     parity JSON report from parity/cli.ts --json",
    "  --doc <path|name>    filter to one doc by absolute path or basename",
    "  --page <n>          filter divergences by page",
    "  --kind <kind>       filter by divergence kind",
    "  --text <text>       substring filter after parity text normalization",
    "  --limit <n>         max rows (default: 25)",
  ].join("\n");

const parseArgs = (argv: string[]): Flags => {
  const flags: Flags = { limit: 25 };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--report") {
      flags.report = argv[++i];
    } else if (arg === "--doc") {
      flags.doc = argv[++i];
    } else if (arg === "--page") {
      flags.page = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--kind") {
      const kind = argv[++i] as DivergenceKind | undefined;
      if (!kind || !KIND_ORDER.includes(kind)) {
        throw new Error(`--kind must be one of: ${KIND_ORDER.join(", ")}`);
      }
      flags.kind = kind;
    } else if (arg === "--text") {
      flags.text = argv[++i];
    } else if (arg === "--limit") {
      flags.limit = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--help" || arg === "-h") {
      console.log(usage());
      process.exit(0);
    } else if (!flags.report) {
      flags.report = arg;
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!flags.report) {
    throw new Error("--report is required");
  }
  if (flags.page !== undefined && (!Number.isInteger(flags.page) || flags.page <= 0)) {
    throw new Error("--page must be a positive integer");
  }
  if (!Number.isInteger(flags.limit) || flags.limit <= 0) {
    throw new Error("--limit must be a positive integer");
  }
  return flags;
};

const readReport = async (file: string): Promise<CorpusReport> => {
  const report = (await Bun.file(path.resolve(file)).json()) as CorpusReport;
  if (!Array.isArray(report.results)) {
    throw new Error(`${file} is not a parity JSON report`);
  }
  if (!report.reference || !isReferenceRendererId(report.reference.id)) {
    throw new Error(`${file} predates explicit reference-renderer metadata; regenerate it`);
  }
  return report;
};

const divergencePage = (divergence: Divergence): number | undefined => {
  switch (divergence.kind) {
    case "page-count":
      return undefined;
    case "pagination":
      return divergence.referencePage;
    case "line-break":
    case "missing-line":
    case "extra-line":
    case "x-drift":
    case "y-drift":
    case "width-drift":
    case "text-mismatch":
      return divergence.page;
  }
};

const divergenceText = (divergence: Divergence): string => {
  switch (divergence.kind) {
    case "page-count":
      return `reference=${divergence.reference} folio=${divergence.folio}`;
    case "pagination":
    case "missing-line":
    case "extra-line":
    case "x-drift":
    case "y-drift":
    case "width-drift":
      return divergence.text;
    case "line-break":
      return `${divergence.referenceTexts.join(" ")} | ${divergence.folioTexts.join(" ")}`;
    case "text-mismatch":
      return `${divergence.referenceText} -> ${divergence.folioText}`;
  }
};

const divergenceQueryText = (divergence: Divergence): string => {
  switch (divergence.kind) {
    case "page-count":
      return divergenceText(divergence);
    case "pagination":
    case "missing-line":
    case "extra-line":
    case "x-drift":
    case "y-drift":
    case "width-drift":
      return divergence.text;
    case "line-break":
      return divergence.referenceTexts.join(" ");
    case "text-mismatch":
      return divergence.referenceText;
  }
};

const divergenceMagnitude = (divergence: Divergence): number | undefined => {
  switch (divergence.kind) {
    case "x-drift":
    case "width-drift":
      return divergence.deltaPt;
    case "y-drift":
      return divergence.residualPt;
    default:
      return undefined;
  }
};

const matchesDoc = (file: string, filter: string | undefined): boolean => {
  if (!filter) {
    return true;
  }
  return file === path.resolve(filter) || path.basename(file) === filter;
};

const shellQuote = (value: string): string => `'${value.replaceAll("'", "'\\''")}'`;

const main = async (): Promise<void> => {
  const flags = parseArgs(process.argv.slice(2));
  const reportPath = path.resolve(flags.report!);
  const report = await readReport(reportPath);
  const textNeedle = flags.text ? normalizeLineText(flags.text).toLocaleLowerCase() : undefined;

  const rows = report.results
    .filter((result) => matchesDoc(result.file, flags.doc))
    .flatMap((result) =>
      result.divergences.map((divergence, index) => ({
        doc: result.file,
        index: index + 1,
        page: divergencePage(divergence),
        kind: divergence.kind,
        text: divergenceText(divergence),
        queryText: divergenceQueryText(divergence),
        magnitudePt: divergenceMagnitude(divergence),
      })),
    )
    .filter((row) => flags.page === undefined || row.page === flags.page)
    .filter((row) => flags.kind === undefined || row.kind === flags.kind)
    .filter(
      (row) =>
        textNeedle === undefined ||
        normalizeLineText(row.text).toLocaleLowerCase().includes(textNeedle),
    )
    .toSorted(
      (a, b) =>
        (a.page ?? Number.MAX_SAFE_INTEGER) - (b.page ?? Number.MAX_SAFE_INTEGER) ||
        KIND_ORDER.indexOf(a.kind) - KIND_ORDER.indexOf(b.kind) ||
        a.index - b.index,
    );

  const selected = rows.slice(0, flags.limit);
  const counts: Partial<Record<DivergenceKind, number>> = {};
  for (const row of rows) {
    counts[row.kind] = (counts[row.kind] ?? 0) + 1;
  }

  console.log(
    JSON.stringify(
      {
        report: reportPath,
        reference: report.reference,
        total: rows.length,
        counts,
        rows: selected.map((row) =>
          Object.assign({}, row, {
            inspectCommand:
              row.page === undefined
                ? undefined
                : `bun parity/inspect.ts --doc ${shellQuote(row.doc)} --reference ${report.reference.id} --page ${row.page} --text ${shellQuote(row.queryText)} --max-pages ${row.page}`,
            sourceTraceCommand: `bun parity/sourceTrace.ts --doc ${shellQuote(row.doc)} --text ${shellQuote(row.queryText)}`,
          }),
        ),
      },
      null,
      2,
    ),
  );
};

try {
  await main();
} catch (error) {
  const err = error instanceof Error ? error : new Error(String(error));
  console.error(`parity divergences failed: ${err.message}`);
  console.error(usage());
  process.exitCode = 2;
}
