/** Report-only L9 over cached, locked corpus files. Does not fetch or update baselines. */
import path from "node:path";
import { Result, TaggedError } from "better-result";
import { normalizeForOps, REVISION_DECISIONS } from "../packages/docx-core/src/ops/documentOps";

import { parseDocx } from "../packages/core/src/docx/parser";
import {
  compareReviewResolution,
  type ReviewOracleOutcome,
} from "../packages/core/src/ops/__tests__/reviewOracle";
import {
  assertOutsideRepository,
  corpusCacheRoot,
  loadCorpusLock,
  sha256Bytes,
} from "./lib/corpus-manifest";

class ReviewReportError extends TaggedError("ReviewReportError")<{ message: string }> {}

type Options = { limit: number; output: string | undefined };
const optionsFrom = (args: readonly string[]): Options => {
  let limit = 100;
  let output: string | undefined;
  for (let index = 0; index < args.length; index += 1) {
    const flag = args.at(index);
    const value = args.at(index + 1);
    if (value === undefined)
      throw new ReviewReportError({ message: "Every argument requires a value." });
    switch (flag) {
      case "--limit":
        limit = Number(value);
        if (!Number.isSafeInteger(limit) || limit < 1) {
          throw new ReviewReportError({ message: "--limit must be a positive integer." });
        }
        break;
      case "--out":
        output = assertOutsideRepository(path.resolve(value));
        break;
      default:
        throw new ReviewReportError({ message: "Supported arguments: --limit N, --out PATH." });
    }
    index += 1;
  }
  return { limit, output };
};

type FileOutcome =
  | ReviewOracleOutcome
  | { type: "missing-cache" }
  | { type: "digest-mismatch" }
  | { type: "file-failed"; errorType: string };
type Finding = { sha256: string; outcome: FileOutcome };

const main = async (): Promise<void> => {
  const options = optionsFrom(process.argv.slice(2));
  const lock = await loadCorpusLock();
  const seen = new Set<string>();
  const files = [...lock.sources]
    .sort((left, right) => left.id.localeCompare(right.id))
    .flatMap((source) =>
      [...source.files]
        .sort((left, right) => left.path.localeCompare(right.path))
        .map((file) => ({
          sha256: file.sha256,
          absolutePath: path.join(corpusCacheRoot(), "sources", source.id, file.path),
        })),
    )
    .filter(({ sha256 }) => {
      if (seen.has(sha256)) return false;
      seen.add(sha256);
      return true;
    });
  const findings: Finding[] = [];
  for (const file of files.slice(0, options.limit)) {
    const local = Bun.file(file.absolutePath);
    if (!(await local.exists())) {
      findings.push({ sha256: file.sha256, outcome: { type: "missing-cache" } });
      continue;
    }
    const read = await Result.tryPromise({
      try: () => local.arrayBuffer(),
      catch: (cause: unknown) => cause,
    });
    if (read.isErr()) {
      findings.push({
        sha256: file.sha256,
        outcome: {
          type: "file-failed",
          errorType: read.error instanceof Error ? read.error.name : "UnknownError",
        },
      });
      continue;
    }
    if (sha256Bytes(new Uint8Array(read.value)) !== file.sha256) {
      findings.push({ sha256: file.sha256, outcome: { type: "digest-mismatch" } });
      continue;
    }
    const parsed = await Result.tryPromise({
      try: async () =>
        normalizeForOps(
          await parseDocx(read.value, {
            preloadFonts: false,
            detectVariables: false,
          }),
        ),
      catch: (cause: unknown) => cause,
    });
    if (parsed.isErr()) {
      findings.push({
        sha256: file.sha256,
        outcome: {
          type: "file-failed",
          errorType: parsed.error instanceof Error ? parsed.error.name : "UnknownError",
        },
      });
      continue;
    }
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const compared = Result.try({
        try: () => compareReviewResolution(parsed.value, decision),
        catch: (cause: unknown) => cause,
      });
      findings.push({
        sha256: file.sha256,
        outcome: compared.isOk()
          ? compared.value
          : {
              type: "file-failed",
              errorType: compared.error instanceof Error ? compared.error.name : "UnknownError",
            },
      });
    }
  }
  const counts: Record<string, number> = {};
  for (const { outcome } of findings) counts[outcome.type] = (counts[outcome.type] ?? 0) + 1;
  const report = `${JSON.stringify(
    {
      schemaVersion: 1,
      oracle: "L9",
      mode: "report-only",
      selectedFiles: Math.min(files.length, options.limit),
      availableFiles: files.length,
      truncated: files.length > options.limit,
      counts,
      findings,
    },
    null,
    2,
  )}\n`;
  if (options.output === undefined) process.stdout.write(report);
  else await Bun.write(options.output, report);
};

// Boundary failures are fatal; measured disagreements leave the CLI successful.
await main();
