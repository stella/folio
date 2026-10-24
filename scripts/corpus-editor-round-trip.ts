/** Measure the editor round-trip family over one shard of tier-1 corpus files. */

import { Result, TaggedError } from "better-result";
import { createHash } from "node:crypto";
import { mkdir } from "node:fs/promises";
import path from "node:path";

import { buildFileList } from "./corpus-gate";
import {
  type CorpusEvidence,
  type CorpusFileResult,
  evidenceOf,
} from "./lib/corpus-census";
import { type CorpusCheckResult, runCorpusChecks } from "./lib/corpus-check";
import type { NotADocxReason } from "./lib/corpus-classify";
import { familyBaselineFromCensus } from "./lib/corpus-family-baseline";
import {
  countedFailures,
  FamilyCensusBuilder,
} from "./lib/corpus-family-census";
import { fileSignatureRows } from "./lib/corpus-file-signatures";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import { CORPUS_TIERS } from "./lib/corpus-manifest";
import {
  assertReportOnlyFilesAreLive,
  loadReportOnlyFiles,
  reportOnlyFileIds,
  reportOnlyFilesDigest,
} from "./lib/corpus-report-only";
import { corpusFileId } from "./lib/corpus-selection";
import { CORPUS_INVARIANTS, type CorpusFailure } from "./lib/corpus-signature";

const INVARIANT_BUDGET_MS = 300_000;
const FILE_BUDGET_MS = 600_000;
const PROGRESS_INTERVAL = 100;

type Shard = { index: number; total: number };

class ScopedCensusError extends TaggedError("ScopedCensusError")<{
  message: string;
}> {}

type IncompleteReason = {
  file: { sourceId: string; path: string; sha256: string };
  reason: string;
};

type FileFailure = {
  file: { sourceId: string; path: string; sha256: string };
  evidence: CorpusEvidence;
  failures: CorpusFailure[];
};

type EditorRoundTripVerdict =
  | {
      kind: "not-a-docx";
      status: "complete";
      editorRun: "not-run";
      evidence: CorpusEvidence;
      parseFailures: [];
      editorFailures: [];
      notADocx: NotADocxReason;
    }
  | {
      kind: "parse-failed";
      status: "complete";
      editorRun: "not-run";
      evidence: CorpusEvidence;
      parseFailures: CorpusFailure[];
      editorFailures: [];
    }
  | {
      kind: "checked";
      status: "complete";
      editorRun: "evaluated";
      evidence: CorpusEvidence;
      parseFailures: CorpusFailure[];
      editorFailures: CorpusFailure[];
    }
  | {
      kind: "checked";
      status: "incomplete";
      editorRun: "evaluated" | "not-run";
      evidence: CorpusEvidence;
      parseFailures: CorpusFailure[];
      editorFailures: CorpusFailure[];
      reason: string;
    }
  | {
      kind: "incomplete";
      status: "incomplete";
      editorRun: "not-run";
      evidence: CorpusEvidence;
      parseFailures: CorpusFailure[];
      editorFailures: CorpusFailure[];
      reason: string;
    };

export const editorRoundTripVerdict = (
  result: CorpusCheckResult,
  listedReportOnly: boolean,
): EditorRoundTripVerdict => {
  if (result.kind === "not-a-docx") {
    return {
      kind: "not-a-docx",
      status: "complete",
      editorRun: "not-run",
      evidence: evidenceOf({ kind: "not-a-docx", reason: result.reason }),
      parseFailures: [],
      editorFailures: [],
      notADocx: result.reason,
    };
  }

  const hasEditorRoundTripTimings = Object.keys(result.timings).some((stage) =>
    stage.startsWith(`${EXTENDED_CORPUS_INVARIANTS.editorRoundTrip}.`),
  );
  const parseFailures = result.failures.filter(
    (failure) => failure.invariant === CORPUS_INVARIANTS.parse,
  );
  let fileResult: CorpusFileResult;
  if (listedReportOnly) {
    fileResult = {
      kind: "report-only",
      failures: result.failures,
      declaredRefusals: result.declaredRefusals,
    };
  } else if (result.truncatedAt === undefined) {
    fileResult = {
      kind: "complete",
      failures: result.failures,
      declaredRefusals: result.declaredRefusals,
    };
  } else {
    fileResult = {
      kind: "truncated",
      failures: result.failures,
      declaredRefusals: result.declaredRefusals,
      stage: result.truncatedAt,
    };
  }

  const evidence = evidenceOf(fileResult);
  const editorFailures = result.failures.filter(
    (failure) =>
      failure.invariant === EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
  );
  if (result.truncatedAt !== undefined) {
    return {
      kind: "checked",
      status: "incomplete",
      editorRun: hasEditorRoundTripTimings ? "evaluated" : "not-run",
      evidence,
      parseFailures,
      editorFailures,
      reason: `Truncated at ${result.truncatedAt}`,
    };
  }
  if (!hasEditorRoundTripTimings && parseFailures.length > 0) {
    return {
      kind: "parse-failed",
      status: "complete",
      editorRun: "not-run",
      evidence,
      parseFailures,
      editorFailures: [],
    };
  }
  if (!hasEditorRoundTripTimings) {
    return {
      kind: "incomplete",
      status: "incomplete",
      editorRun: "not-run",
      evidence,
      parseFailures,
      editorFailures,
      reason: "Editor round-trip invariant did not run",
    };
  }
  return {
    kind: "checked",
    status: "complete",
    editorRun: "evaluated",
    evidence,
    parseFailures,
    editorFailures,
  };
};

type MeasuredFile =
  | { kind: "sha-mismatch"; actualSha256: string }
  | { kind: "checked"; result: CorpusCheckResult };

const parseOptions = (args: readonly string[]) => {
  let shardValue: string | undefined;
  let outPath: string | undefined;
  for (let index = 0; index < args.length; index += 1) {
    const flag = args.at(index);
    const value = args.at(index + 1);
    if (flag !== "--shard" && flag !== "--out") {
      throw new ScopedCensusError({
        message: `Unexpected argument: ${flag ?? "<missing>"}`,
      });
    }
    if (value === undefined || value.startsWith("--")) {
      throw new ScopedCensusError({ message: `${flag} needs a value` });
    }
    if (value.length === 0) {
      throw new ScopedCensusError({ message: `${flag} needs a value` });
    }
    if (flag === "--shard") {
      if (shardValue !== undefined) {
        throw new ScopedCensusError({ message: "--shard may appear once" });
      }
      shardValue = value;
    } else {
      if (outPath !== undefined) {
        throw new ScopedCensusError({ message: "--out may appear once" });
      }
      outPath = value;
    }
    index += 1;
  }
  if (shardValue === undefined || outPath === undefined) {
    throw new ScopedCensusError({
      message:
        "Usage: bun scripts/corpus-editor-round-trip.ts --shard k/n --out FILE",
    });
  }
  const match = /^(\d+)\/(\d+)$/u.exec(shardValue);
  const index = Number(match?.[1]);
  const total = Number(match?.[2]);
  if (
    !Number.isInteger(index) ||
    !Number.isInteger(total) ||
    total < 1 ||
    index < 1 ||
    index > total
  ) {
    throw new ScopedCensusError({
      message: `--shard expects k/n with 1 <= k <= n, got ${shardValue}`,
    });
  }
  return { shard: { index, total }, outPath };
};

const currentHead = (): string => {
  const resolved = Bun.spawnSync(["git", "rev-parse", "HEAD"], {
    stdout: "pipe",
    stderr: "pipe",
  });
  const head = resolved.stdout.toString().trim();
  if (resolved.exitCode !== 0 || head.length === 0) {
    throw new ScopedCensusError({
      message: "Could not resolve the measured git head",
    });
  }
  const fromEnvironment = Bun.env["GITHUB_SHA"]?.trim();
  if (fromEnvironment && fromEnvironment !== head) {
    throw new ScopedCensusError({
      message: `GITHUB_SHA ${fromEnvironment} does not match checked-out HEAD ${head}`,
    });
  }
  const tracked = Bun.spawnSync(["git", "status", "--porcelain", "-uno"], {
    stdout: "pipe",
    stderr: "pipe",
  });
  if (tracked.exitCode !== 0 || tracked.stdout.toString().trim() !== "") {
    throw new ScopedCensusError({
      message: "A scoped census requires a clean tracked worktree",
    });
  }
  return head;
};

const fileId = (entry: {
  sourceId: string;
  relativePath: string;
  sha256: string;
}) => ({
  sourceId: entry.sourceId,
  path: entry.relativePath,
  sha256: entry.sha256,
});

const run = async (shard: Shard, outPath: string): Promise<void> => {
  const head = currentHead();
  const tiers = [CORPUS_TIERS.permissive];
  const { entries, lockDigest, lock } = await buildFileList(tiers);
  const reportOnly = await loadReportOnlyFiles();
  assertReportOnlyFilesAreLive(reportOnly, lock);
  const reportOnlyIds = reportOnlyFileIds(reportOnly);
  const reportOnlyDigest = reportOnlyFilesDigest(reportOnly);
  const assigned = entries.filter(
    (_, position) => position % shard.total === shard.index - 1,
  );
  if (assigned.length === 0) {
    throw new ScopedCensusError({
      message: "The selected tier-1 corpus shard is empty",
    });
  }

  const family = new FamilyCensusBuilder(lockDigest, reportOnlyDigest);
  const incomplete: IncompleteReason[] = [];
  const notADocx: { file: ReturnType<typeof fileId>; reason: string }[] = [];
  const failures: FileFailure[] = [];
  const parseFailures: FileFailure[] = [];
  let duplicates = 0;
  let checked = 0;
  let editorEvaluatedFiles = 0;

  for (const entry of assigned) {
    if (entry.duplicateOf !== null) {
      duplicates += 1;
      continue;
    }

    const file = fileId(entry);
    const id = corpusFileId(entry);
    const measured = await Result.tryPromise({
      try: async (): Promise<MeasuredFile> => {
        const bytes = new Uint8Array(
          await Bun.file(entry.absolutePath).arrayBuffer(),
        );
        const actualSha256 = createHash("sha256").update(bytes).digest("hex");
        if (actualSha256 !== entry.sha256) {
          return { kind: "sha-mismatch", actualSha256 };
        }
        return {
          kind: "checked",
          result: await runCorpusChecks(bytes, {
            invariantBudgetMs: INVARIANT_BUDGET_MS,
            fileBudgetMs: FILE_BUDGET_MS,
            only: new Set([EXTENDED_CORPUS_INVARIANTS.editorRoundTrip]),
          }),
        };
      },
      catch: (cause: unknown) => cause,
    });
    if (measured.isErr()) {
      incomplete.push({
        file,
        reason:
          measured.error instanceof Error
            ? measured.error.message
            : String(measured.error),
      });
      continue;
    }

    if (measured.value.kind === "sha-mismatch") {
      incomplete.push({
        file,
        reason: `SHA-256 mismatch: locked ${entry.sha256}, received ${measured.value.actualSha256}`,
      });
      continue;
    }

    const { result } = measured.value;
    if (result.kind === "not-a-docx") {
      notADocx.push({ file, reason: result.reason });
      continue;
    }

    const verdict = editorRoundTripVerdict(result, reportOnlyIds.has(id));
    checked += 1;
    if (verdict.editorRun === "evaluated") editorEvaluatedFiles += 1;
    if (verdict.parseFailures.length > 0) {
      parseFailures.push({
        file,
        evidence: verdict.evidence,
        failures: verdict.parseFailures,
      });
    }
    if (verdict.editorFailures.length > 0) {
      failures.push({
        file,
        evidence: verdict.evidence,
        failures: verdict.editorFailures,
      });
    }
    if (verdict.status === "incomplete") {
      incomplete.push({ file, reason: verdict.reason });
    }

    family.add({
      file,
      bytes: result.cost.bytes,
      parseMs: result.cost.parseMs,
      peakRssBytes: result.cost.peakRssBytes,
      referenceMs: result.cost.referenceMs,
      producer: result.producer,
      failures: result.failures,
      declaredRefusals: result.declaredRefusals,
      timings: result.timings,
      evidence: verdict.evidence,
    });

    if (checked > 0 && checked % PROGRESS_INTERVAL === 0) {
      process.stderr.write(
        `  ${checked} checked files in shard ${shard.index}/${shard.total}\n`,
      );
    }
  }

  const familyCensus = family.build();
  if (currentHead() !== head) {
    throw new ScopedCensusError({ message: "The measured git head changed" });
  }
  const status = incomplete.length === 0 ? "complete" : "incomplete";
  const artifact = {
    schemaVersion: 1,
    scope: "editor-round-trip",
    head,
    lockDigest: familyCensus.lockDigest,
    reportOnlyDigest: familyCensus.reportOnlyDigest,
    status,
    shard,
    assignedFiles: assigned.length,
    checkedFiles: checked,
    editorEvaluatedFiles,
    duplicates,
    notADocx,
    incomplete,
    parseFailures,
    metadata: {
      denominator: "all-checked-results-including-parse-refusals",
      editorEvaluation: "results-with-editor-round-trip-timings",
      completion: "parse-refusals-count-as-complete-editor-scoped-outcomes",
    },
    declaredRefusals: familyCensus.declaredRefusals,
    failures,
    fileSignatures: failures.flatMap(({ file, evidence, failures: findings }) =>
      fileSignatureRows(file, countedFailures(findings, evidence)),
    ),
    baseline: familyBaselineFromCensus(
      familyCensus,
      EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
    ),
  } as const;

  await mkdir(path.dirname(path.resolve(outPath)), { recursive: true });
  await Bun.write(outPath, `${JSON.stringify(artifact, null, 2)}\n`);
  process.stdout.write(
    `Editor round-trip census ${shard.index}/${shard.total}: ${status}; ${checked} checked, ${duplicates} duplicates, ${notADocx.length} not-a-docx, ${incomplete.length} incomplete -> ${outPath}\n`,
  );
  if (status === "incomplete") process.exitCode = 1;
};

if (import.meta.main) {
  try {
    const { shard, outPath } = parseOptions(process.argv.slice(2));
    await run(shard, outPath);
  } catch (cause) {
    process.stderr.write(
      `${cause instanceof Error ? cause.message : String(cause)}\n`,
    );
    process.exitCode = 1;
  }
}
