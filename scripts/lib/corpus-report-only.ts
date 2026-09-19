/**
 * The files whose findings are measured but never gated.
 *
 * A per-file budget can stop a run part-way, and whether it does depends on how
 * busy the runner was. So "which files contributed gating evidence" was a
 * function of wall-clock time: the same commit compared a different set of
 * files from one night to the next, and a signature appeared or vanished with
 * no code change behind it.
 *
 * This list makes that set committed data. A listed file is still fetched, run
 * and measured: its timings stay in the report-only `performance` family and in
 * the census, so its cost remains visible. It simply never contributes a gating
 * finding, whether it finished or not. Nothing else may leave the gating set,
 * so an unlisted file that stops at a budget degrades the whole run rather than
 * quietly shrinking what is compared.
 *
 * Every entry names the corpus file by content as well as by path, and a dead
 * entry fails the gate: a repin that drops or rewrites a listed file must be
 * noticed, not carried as a silent exemption.
 */

import { createHash } from "node:crypto";
import path from "node:path";

import { TaggedError } from "better-result";

import { fileIdOf } from "./corpus-census";
import { CORPUS_DIRECTORY, type CorpusLock, isSha256 } from "./corpus-manifest";

export class ReportOnlyFilesError extends TaggedError("ReportOnlyFilesError")<{
  message: string;
}> {}

export const REPORT_ONLY_FILES_PATH = path.join(CORPUS_DIRECTORY, "report-only-files.json");

export type ReportOnlyFile = {
  sourceId: string;
  path: string;
  sha256: string;
  /** Why this file cannot be gated on, in the reviewer's own words. */
  reason: string;
};

export type ReportOnlyFiles = {
  schemaVersion: 1;
  files: ReportOnlyFile[];
};

const ENTRY_KEYS = new Set(["path", "reason", "sha256", "sourceId"]);
const LIST_KEYS = new Set(["files", "schemaVersion"]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const validateEntry = (value: unknown, index: number, issues: string[]): void => {
  const location = `files[${index}]`;
  if (!isRecord(value)) {
    issues.push(`${location}: expected an object`);
    return;
  }
  for (const key of Object.keys(value)) {
    if (!ENTRY_KEYS.has(key)) {
      issues.push(`${location}.${key}: unknown field`);
    }
  }
  for (const key of ["sourceId", "path", "reason"]) {
    const field = value[key];
    if (typeof field !== "string" || field.length === 0) {
      issues.push(`${location}.${key}: expected a non-empty string`);
    }
  }
  const { sha256 } = value;
  if (typeof sha256 !== "string" || !isSha256(sha256)) {
    issues.push(`${location}.sha256: expected a lowercase SHA-256`);
  }
};

export const validateReportOnlyFiles = (value: unknown): string[] => {
  const issues: string[] = [];
  if (!isRecord(value)) {
    return ["report-only files: expected an object"];
  }
  for (const key of Object.keys(value)) {
    if (!LIST_KEYS.has(key)) {
      issues.push(`${key}: unknown field`);
    }
  }
  if (value["schemaVersion"] !== 1) {
    issues.push("schemaVersion: expected 1");
  }
  const files = value["files"];
  if (!Array.isArray(files)) {
    issues.push("files: expected an array");
    return issues;
  }
  for (const [index, entry] of files.entries()) {
    validateEntry(entry, index, issues);
  }
  // Sorted and unique: the file is read by humans deciding whether an exemption
  // is still earned, and the digest below has to be a function of the set.
  const ids = files.flatMap((entry) =>
    isRecord(entry) && typeof entry["sourceId"] === "string" && typeof entry["path"] === "string"
      ? [fileIdOf({ sourceId: entry["sourceId"], path: entry["path"] })]
      : [],
  );
  if (new Set(ids).size !== ids.length) {
    issues.push("files: the same file is listed twice");
  }
  if (ids.some((id, index) => id < (ids[index - 1] ?? ""))) {
    issues.push("files: entries must be sorted by source id, then path");
  }
  return issues;
};

/**
 * The digest a baseline records, over the set alone.
 *
 * Only the identity of the listed files changes what is compared, so rewording
 * a reason must not invalidate a night's measurement.
 */
export const reportOnlyFilesDigest = (list: ReportOnlyFiles): string => {
  const hash = createHash("sha256");
  for (const file of list.files) {
    hash.update(`${file.sourceId}\u0000${file.path}\u0000${file.sha256}\n`, "utf8");
  }
  return hash.digest("hex");
};

export const reportOnlyFileIds = (list: ReportOnlyFiles): ReadonlySet<string> =>
  new Set(list.files.map((file) => fileIdOf(file)));

/**
 * Entries the corpus no longer carries, or carries with other bytes.
 *
 * An exemption is granted for a file, not for a path: a repin that replaces the
 * content must be reviewed again rather than inheriting the old decision.
 */
export const deadReportOnlyEntries = (list: ReportOnlyFiles, lock: CorpusLock): string[] => {
  const locked = new Map(
    lock.sources.flatMap((source) =>
      source.files.map(
        (file) => [fileIdOf({ sourceId: source.id, path: file.path }), file.sha256] as const,
      ),
    ),
  );
  return list.files.flatMap((file) => {
    const id = fileIdOf(file);
    const sha256 = locked.get(id);
    if (sha256 === undefined) {
      return [`${id}: no such file in corpus/sources.lock.json`];
    }
    return sha256 === file.sha256
      ? []
      : [
          `${id}: the lock records ${sha256.slice(0, 12)}, the entry claims ${file.sha256.slice(0, 12)}`,
        ];
  });
};

export const loadReportOnlyFiles = async (): Promise<ReportOnlyFiles> => {
  const file = Bun.file(REPORT_ONLY_FILES_PATH);
  if (!(await file.exists())) {
    // Absent, every listed file would gate again and the gating set would go
    // back to depending on how fast the runner was.
    throw new ReportOnlyFilesError({
      message:
        "corpus/report-only-files.json is missing. It is committed; restore it rather than running without it.",
    });
  }
  const value = (await file.json()) as unknown;
  const issues = validateReportOnlyFiles(value);
  if (issues.length > 0) {
    throw new ReportOnlyFilesError({
      message: `corpus/report-only-files.json is invalid:\n${issues.map((issue) => `- ${issue}`).join("\n")}`,
    });
  }
  return value as ReportOnlyFiles;
};

export const assertReportOnlyFilesAreLive = (list: ReportOnlyFiles, lock: CorpusLock): void => {
  const dead = deadReportOnlyEntries(list, lock);
  if (dead.length > 0) {
    throw new ReportOnlyFilesError({
      message: `corpus/report-only-files.json lists files the corpus does not carry:\n${dead
        .map((issue) => `- ${issue}`)
        .join("\n")}`,
    });
  }
};
