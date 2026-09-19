/**
 * The public DOCX corpus manifest and its per-file lock.
 *
 * The corpus is third-party content: nothing it contains is ever committed.
 * `corpus/sources.json` names public repositories, a pinned commit, the
 * sub-paths to take and the licence each set was reviewed under;
 * `corpus/sources.lock.json` records only a relative path, a SHA-256 and a byte
 * count per file, which is enough to verify a cache directory without carrying
 * any of its bytes.
 */

import { createHash } from "node:crypto";
import { homedir } from "node:os";
import path from "node:path";

import { TaggedError } from "better-result";

export class CorpusManifestError extends TaggedError("CorpusManifestError")<{
  message: string;
  cause?: unknown;
}> {}

type CorpusLicense = {
  spdxId: string;
  auditStatus: "needs-review" | "reviewed";
  noticeUrl: string;
  redistribution: "cache-only";
};

export type CorpusSource = {
  id: string;
  title: string;
  publisher: string;
  repository: string;
  commit: string;
  tree: string;
  paths: string[];
  license: CorpusLicense;
};

export type CorpusManifest = {
  schemaVersion: 1;
  fileLimit: number;
  maxFileBytes: number;
  sources: CorpusSource[];
};

export type CorpusLockEntry = {
  path: string;
  sha256: string;
  bytes: number;
};

export type CorpusLockSource = {
  id: string;
  commit: string;
  files: CorpusLockEntry[];
};

export type CorpusLock = {
  schemaVersion: 1;
  manifestDigest: string;
  fileCount: number;
  totalBytes: number;
  sources: CorpusLockSource[];
};

export const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..", "..");
export const CORPUS_DIRECTORY = path.join(REPOSITORY_ROOT, "corpus");
export const MANIFEST_PATH = path.join(CORPUS_DIRECTORY, "sources.json");
export const LOCK_PATH = path.join(CORPUS_DIRECTORY, "sources.lock.json");
export const BASELINE_PATH = path.join(CORPUS_DIRECTORY, "baseline.json");

const SOURCE_ID_RE = /^[a-z0-9]+(?:-[a-z0-9]+)*$/u;
const SHA256_RE = /^[a-f0-9]{64}$/u;
const GIT_OBJECT_ID_RE = /^[a-f0-9]{40}$/u;
const GITHUB_REPOSITORY_RE = /^https:\/\/github\.com\/[\w.-]+\/[\w.-]+$/u;
const HTTPS_URL_RE = /^https:\/\//u;
const SPDX_ID_RE = /^[\w.-]+(?: (?:OR|AND|WITH) [\w.-]+)*$/u;
/** A sparse-checkout pattern: rooted, no `..`, and restricted to `.docx`. */
const SOURCE_PATH_RE = /^\/[\w./*-]*\*\.docx$/u;

const AUDIT_STATUSES = new Set<unknown>(["needs-review", "reviewed"]);
const REDISTRIBUTION_POLICIES = new Set<unknown>(["cache-only"]);
const LICENSE_KEYS = new Set(["auditStatus", "noticeUrl", "redistribution", "spdxId"]);
const SOURCE_KEYS = new Set([
  "commit",
  "id",
  "license",
  "paths",
  "publisher",
  "repository",
  "title",
  "tree",
]);
const MANIFEST_KEYS = new Set(["fileLimit", "maxFileBytes", "schemaVersion", "sources"]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const compareStrings = (left: string, right: string): number => {
  if (left < right) {
    return -1;
  }
  return left > right ? 1 : 0;
};

const validateExactKeys = (
  record: Record<string, unknown>,
  allowedKeys: ReadonlySet<string>,
  location: string,
  issues: string[],
): void => {
  for (const key of Object.keys(record)) {
    if (!allowedKeys.has(key)) {
      issues.push(`${location}.${key}: unknown field`);
    }
  }
};

const validateRequiredString = (
  record: Record<string, unknown>,
  key: string,
  location: string,
  issues: string[],
): void => {
  const value = record[key];
  if (typeof value !== "string" || value.length === 0) {
    issues.push(`${location}.${key}: expected a non-empty string`);
  }
};

const validateLicense = (value: unknown, location: string, issues: string[]): void => {
  if (!isRecord(value)) {
    issues.push(`${location}: expected an object`);
    return;
  }
  validateExactKeys(value, LICENSE_KEYS, location, issues);
  validateRequiredString(value, "noticeUrl", location, issues);
  const { noticeUrl, redistribution, spdxId } = value;
  if (typeof noticeUrl === "string" && !HTTPS_URL_RE.test(noticeUrl)) {
    issues.push(`${location}.noticeUrl: expected an HTTPS URL`);
  }
  if (typeof spdxId !== "string" || !SPDX_ID_RE.test(spdxId)) {
    issues.push(`${location}.spdxId: expected an SPDX licence expression`);
  }
  if (!AUDIT_STATUSES.has(value["auditStatus"])) {
    issues.push(`${location}.auditStatus: expected needs-review or reviewed`);
  }
  if (!REDISTRIBUTION_POLICIES.has(redistribution)) {
    issues.push(`${location}.redistribution: corpus content is cache-only`);
  }
};

const validatePaths = (value: unknown, location: string, issues: string[]): void => {
  if (!Array.isArray(value) || value.length === 0) {
    issues.push(`${location}: expected a non-empty array`);
    return;
  }
  for (const [index, pattern] of value.entries()) {
    if (typeof pattern !== "string" || !SOURCE_PATH_RE.test(pattern)) {
      issues.push(`${location}[${index}]: expected a rooted \`*.docx\` sparse-checkout pattern`);
      continue;
    }
    if (pattern.split("/").includes("..")) {
      issues.push(`${location}[${index}]: must not traverse upwards`);
    }
  }
};

const validateSource = (value: unknown, index: number, issues: string[]): void => {
  const location = `sources[${index}]`;
  if (!isRecord(value)) {
    issues.push(`${location}: expected an object`);
    return;
  }
  validateExactKeys(value, SOURCE_KEYS, location, issues);
  for (const key of ["id", "publisher", "repository", "title"]) {
    validateRequiredString(value, key, location, issues);
  }
  const { id, repository } = value;
  if (typeof id === "string" && !SOURCE_ID_RE.test(id)) {
    issues.push(`${location}.id: expected kebab-case`);
  }
  if (typeof repository === "string" && !GITHUB_REPOSITORY_RE.test(repository)) {
    issues.push(`${location}.repository: expected a GitHub HTTPS repository URL`);
  }
  for (const key of ["commit", "tree"]) {
    const objectId = value[key];
    if (typeof objectId !== "string" || !GIT_OBJECT_ID_RE.test(objectId)) {
      issues.push(`${location}.${key}: expected a lowercase Git object ID`);
    }
  }
  validatePaths(value["paths"], `${location}.paths`, issues);
  validateLicense(value["license"], `${location}.license`, issues);
};

export const validateCorpusManifest = (value: unknown): string[] => {
  const issues: string[] = [];
  if (!isRecord(value)) {
    return ["manifest: expected an object"];
  }
  validateExactKeys(value, MANIFEST_KEYS, "manifest", issues);
  if (value["schemaVersion"] !== 1) {
    issues.push("manifest.schemaVersion: expected 1");
  }
  for (const key of ["fileLimit", "maxFileBytes"]) {
    const limit = value[key];
    if (typeof limit !== "number" || !Number.isSafeInteger(limit) || limit <= 0) {
      issues.push(`manifest.${key}: expected a positive safe integer`);
    }
  }
  const sources = value["sources"];
  if (!Array.isArray(sources) || sources.length === 0) {
    issues.push("manifest.sources: expected a non-empty array");
    return issues;
  }
  for (const [index, source] of sources.entries()) {
    validateSource(source, index, issues);
  }
  const ids = sources.flatMap((source) =>
    isRecord(source) && typeof source["id"] === "string" ? [source["id"]] : [],
  );
  if (new Set(ids).size !== ids.length) {
    issues.push("manifest.sources: duplicate source ids");
  }
  if (ids.toSorted(compareStrings).some((id, index) => id !== ids[index])) {
    issues.push("manifest.sources: sources must be sorted by id");
  }
  return issues;
};

const isCorpusManifest = (value: unknown): value is CorpusManifest =>
  validateCorpusManifest(value).length === 0;

const parseJson = (text: string, what: string): unknown => {
  try {
    return JSON.parse(text);
  } catch (cause) {
    throw new CorpusManifestError({ message: `${what} is not valid JSON`, cause });
  }
};

/**
 * The manifest's content, independent of how the file is spelled.
 *
 * The digest below binds a lock to the manifest that produced it, so it must
 * survive a reformat: a formatter run is not a corpus change and must not
 * invalidate a hundred megabytes of cache.
 */
const canonicalJson = (value: unknown): string => {
  if (Array.isArray(value)) {
    return `[${value.map(canonicalJson).join(",")}]`;
  }
  if (typeof value === "object" && value !== null) {
    const entries = Object.entries(value)
      .toSorted(([left], [right]) => compareStrings(left, right))
      .map(([key, entry]) => `${JSON.stringify(key)}:${canonicalJson(entry)}`);
    return `{${entries.join(",")}}`;
  }
  return JSON.stringify(value) ?? "null";
};

export type LoadedCorpusManifest = {
  manifest: CorpusManifest;
  /** Binds a lock to the manifest that produced it. */
  digest: string;
};

export const loadCorpusManifest = async (): Promise<LoadedCorpusManifest> => {
  const text = await Bun.file(MANIFEST_PATH).text();
  const value = parseJson(text, "corpus/sources.json");
  const issues = validateCorpusManifest(value);
  if (issues.length > 0) {
    throw new CorpusManifestError({
      message: `corpus/sources.json is invalid:\n${issues.map((issue) => `- ${issue}`).join("\n")}`,
    });
  }
  if (!isCorpusManifest(value)) {
    throw new CorpusManifestError({ message: "corpus/sources.json validation failed" });
  }
  const unreviewed = value.sources.filter((source) => source.license.auditStatus !== "reviewed");
  if (unreviewed.length > 0) {
    throw new CorpusManifestError({
      message: `Licence review is pending for: ${unreviewed.map((source) => source.id).join(", ")}`,
    });
  }
  return { manifest: value, digest: sha256Text(canonicalJson(value)) };
};

export const loadCorpusLock = async (): Promise<CorpusLock> => {
  const file = Bun.file(LOCK_PATH);
  if (!(await file.exists())) {
    throw new CorpusManifestError({
      message: "corpus/sources.lock.json is missing. Run `bun run corpus:lock`.",
    });
  }
  const value = parseJson(await file.text(), "corpus/sources.lock.json");
  if (!isRecord(value) || value["schemaVersion"] !== 1 || !Array.isArray(value["sources"])) {
    throw new CorpusManifestError({ message: "corpus/sources.lock.json has an unknown shape" });
  }
  // SAFETY: the shape check above plus `corpusLockDigest`'s field reads are the
  // only contract this lock has; it is generated by `corpus-sources.ts lock`.
  return value as unknown as CorpusLock;
};

export const sha256Text = (text: string): string =>
  createHash("sha256").update(text, "utf8").digest("hex");

export const sha256Bytes = (bytes: Uint8Array): string =>
  createHash("sha256").update(bytes).digest("hex");

export const isSha256 = (value: string): boolean => SHA256_RE.test(value);

/**
 * A digest over the exact corpus a run saw.
 *
 * The baseline records it so that changing the manifest, repinning a commit or
 * gaining a file forces an explicit baseline refresh rather than silently
 * shifting the failure counts the ratchet compares against.
 */
export const corpusLockDigest = (lock: CorpusLock): string => {
  const hash = createHash("sha256");
  for (const source of lock.sources) {
    for (const file of source.files) {
      hash.update(`${source.id}\u0000${file.path}\u0000${file.sha256}\n`, "utf8");
    }
  }
  return hash.digest("hex");
};

const DEFAULT_CACHE_DIRECTORY = path.join(homedir(), ".cache", "folio-corpus");

/**
 * Refuse a destination inside the repository.
 *
 * Corpus content and anything derived from it (a minimised reproduction is
 * still third-party content) stays out of the working tree, where it cannot be
 * committed by accident.
 */
export const assertOutsideRepository = (directory: string): string => {
  const resolved = path.resolve(directory);
  if (resolved === REPOSITORY_ROOT || resolved.startsWith(`${REPOSITORY_ROOT}${path.sep}`)) {
    throw new CorpusManifestError({
      message: `${resolved} is inside the repository; corpus content is never committed.`,
    });
  }
  return resolved;
};

/**
 * Where fetched corpus content lives: outside the repository, always.
 *
 * `FOLIO_CORPUS_CACHE` overrides it for CI, which restores the same directory
 * from its own cache keyed by the lock digest.
 */
export const corpusCacheRoot = (): string => {
  const override = Bun.env["FOLIO_CORPUS_CACHE"];
  const root = override === undefined || override.length === 0 ? DEFAULT_CACHE_DIRECTORY : override;
  return assertOutsideRepository(root);
};

export const sourceCheckoutPath = (sourceId: string): string =>
  path.join(corpusCacheRoot(), "sources", sourceId);

export const writeJsonFile = async (filePath: string, value: unknown): Promise<void> => {
  await Bun.write(filePath, `${JSON.stringify(value, null, 2)}\n`);
};
