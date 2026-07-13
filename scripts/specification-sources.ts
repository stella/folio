import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { mkdir, rm, stat } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

type LicenseAuditStatus = "needs-review" | "reviewed";
type RedistributionPolicy = "cache-only" | "license-compliant" | "reference-only";
type GeneratedOutputPolicy = "implementation-facts-only" | "license-compliant" | "review-required";

type SourceLicense = {
  auditStatus: LicenseAuditStatus;
  notice: string;
  noticeUrl: string;
  redistribution: RedistributionPolicy;
  generatedOutput: GeneratedOutputPolicy;
};

type SpecificationSourceBase = {
  id: string;
  title: string;
  publisher: string;
  version: string;
  profiles: string[];
  cachePath: string;
  license: SourceLicense;
};

type ArtifactSpecificationSource = SpecificationSourceBase & {
  sourceType: "artifact";
  url: string;
  sha256: string;
  bytes: number;
};

type GitSpecificationSource = SpecificationSourceBase & {
  sourceType: "git";
  repository: string;
  commit: string;
  tree: string;
};

type SpecificationSource = ArtifactSpecificationSource | GitSpecificationSource;

type SpecificationSourceManifest = {
  schemaVersion: 1;
  sources: SpecificationSource[];
};

class SpecificationSourceError extends TaggedError("SpecificationSourceError")<{
  message: string;
  cause?: unknown;
}>() {}

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const MANIFEST_PATH = path.join(REPOSITORY_ROOT, "specifications", "sources.json");
const CACHE_ROOT = path.join(REPOSITORY_ROOT, ".cache", "specifications");
const SOURCE_FETCH_TIMEOUT_MS = 120_000;
const SOURCE_ID_RE = /^[a-z0-9]+(?:-[a-z0-9]+)*$/u;
const SHA256_RE = /^[a-f0-9]{64}$/u;
const GIT_OBJECT_ID_RE = /^[a-f0-9]{40}$/u;
const HTTPS_URL_RE = /^https:\/\//u;
const CACHE_PATH_RE = /^\.cache\/specifications\/[a-z0-9][a-z0-9./-]*$/u;
const ARTIFACT_KEYS = new Set([
  "bytes",
  "cachePath",
  "id",
  "license",
  "profiles",
  "publisher",
  "sha256",
  "sourceType",
  "title",
  "url",
  "version",
]);
const GIT_KEYS = new Set([
  "cachePath",
  "commit",
  "id",
  "license",
  "profiles",
  "publisher",
  "repository",
  "sourceType",
  "title",
  "tree",
  "version",
]);
const LICENSE_KEYS = new Set([
  "auditStatus",
  "generatedOutput",
  "notice",
  "noticeUrl",
  "redistribution",
]);
const LICENSE_AUDIT_STATUSES = new Set<unknown>(["needs-review", "reviewed"]);
const REDISTRIBUTION_POLICIES = new Set<unknown>([
  "cache-only",
  "license-compliant",
  "reference-only",
]);
const GENERATED_OUTPUT_POLICIES = new Set<unknown>([
  "implementation-facts-only",
  "license-compliant",
  "review-required",
]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const compareStrings = (left: string, right: string): number => {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
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

const validateProfiles = (value: unknown, location: string, issues: string[]): void => {
  if (!Array.isArray(value) || value.length === 0) {
    issues.push(`${location}: expected a non-empty array`);
    return;
  }
  if (!value.every((profile) => typeof profile === "string" && SOURCE_ID_RE.test(profile))) {
    issues.push(`${location}: expected kebab-case profile names`);
    return;
  }
  const profiles = value.filter((profile): profile is string => typeof profile === "string");
  if (new Set(profiles).size !== profiles.length) {
    issues.push(`${location}: duplicate profiles`);
  }
  if (profiles.toSorted(compareStrings).some((profile, index) => profile !== profiles[index])) {
    issues.push(`${location}: profiles must be sorted`);
  }
};

const validateLicense = (value: unknown, location: string, issues: string[]): void => {
  if (!isRecord(value)) {
    issues.push(`${location}: expected an object`);
    return;
  }
  validateExactKeys(value, LICENSE_KEYS, location, issues);
  validateRequiredString(value, "notice", location, issues);
  validateRequiredString(value, "noticeUrl", location, issues);
  if (!LICENSE_AUDIT_STATUSES.has(value["auditStatus"])) {
    issues.push(`${location}.auditStatus: unsupported policy`);
  }
  if (!REDISTRIBUTION_POLICIES.has(value["redistribution"])) {
    issues.push(`${location}.redistribution: unsupported policy`);
  }
  if (!GENERATED_OUTPUT_POLICIES.has(value["generatedOutput"])) {
    issues.push(`${location}.generatedOutput: unsupported policy`);
  }
  const noticeUrl = value["noticeUrl"];
  if (typeof noticeUrl === "string" && !HTTPS_URL_RE.test(noticeUrl)) {
    issues.push(`${location}.noticeUrl: expected an HTTPS URL`);
  }
};

const validateSource = (value: unknown, index: number, issues: string[]): void => {
  const location = `sources[${index}]`;
  if (!isRecord(value)) {
    issues.push(`${location}: expected an object`);
    return;
  }
  const sourceType = value["sourceType"];
  if (sourceType !== "artifact" && sourceType !== "git") {
    issues.push(`${location}.sourceType: expected artifact or git`);
    return;
  }
  validateExactKeys(value, sourceType === "artifact" ? ARTIFACT_KEYS : GIT_KEYS, location, issues);
  for (const key of ["cachePath", "id", "publisher", "title", "version"]) {
    validateRequiredString(value, key, location, issues);
  }
  const id = value["id"];
  if (typeof id === "string" && !SOURCE_ID_RE.test(id)) {
    issues.push(`${location}.id: expected kebab-case`);
  }
  const cachePath = value["cachePath"];
  if (
    typeof cachePath === "string" &&
    (!CACHE_PATH_RE.test(cachePath) || cachePath.split("/").includes(".."))
  ) {
    issues.push(`${location}.cachePath: expected a path under .cache/specifications`);
  }
  validateProfiles(value["profiles"], `${location}.profiles`, issues);
  validateLicense(value["license"], `${location}.license`, issues);

  if (sourceType === "artifact") {
    validateRequiredString(value, "url", location, issues);
    const url = value["url"];
    if (typeof url === "string" && !HTTPS_URL_RE.test(url)) {
      issues.push(`${location}.url: expected an HTTPS URL`);
    }
    const sha256 = value["sha256"];
    if (typeof sha256 !== "string" || !SHA256_RE.test(sha256)) {
      issues.push(`${location}.sha256: expected a lowercase SHA-256 digest`);
    }
    const bytes = value["bytes"];
    if (typeof bytes !== "number" || !Number.isSafeInteger(bytes) || bytes <= 0) {
      issues.push(`${location}.bytes: expected a positive safe integer`);
    }
    return;
  }

  validateRequiredString(value, "repository", location, issues);
  const repository = value["repository"];
  if (typeof repository === "string" && !repository.startsWith("https://github.com/")) {
    issues.push(`${location}.repository: expected a GitHub HTTPS URL`);
  }
  for (const key of ["commit", "tree"]) {
    const objectId = value[key];
    if (typeof objectId !== "string" || !GIT_OBJECT_ID_RE.test(objectId)) {
      issues.push(`${location}.${key}: expected a lowercase Git object ID`);
    }
  }
};

export const validateSpecificationSourceManifest = (value: unknown): string[] => {
  const issues: string[] = [];
  if (!isRecord(value)) {
    return ["manifest: expected an object"];
  }
  validateExactKeys(value, new Set(["schemaVersion", "sources"]), "manifest", issues);
  if (value["schemaVersion"] !== 1) {
    issues.push("manifest.schemaVersion: expected 1");
  }
  const sources = value["sources"];
  if (!Array.isArray(sources) || sources.length === 0) {
    issues.push("manifest.sources: expected a non-empty array");
    return issues;
  }
  for (const [index, source] of sources.entries()) {
    validateSource(source, index, issues);
  }
  const ids = sources.flatMap((source) => {
    if (!isRecord(source) || typeof source["id"] !== "string") {
      return [];
    }
    return [source["id"]];
  });
  if (new Set(ids).size !== ids.length) {
    issues.push("manifest.sources: duplicate source ids");
  }
  if (ids.toSorted(compareStrings).some((id, index) => id !== ids[index])) {
    issues.push("manifest.sources: sources must be sorted by id");
  }
  return issues;
};

const isSpecificationSourceManifest = (value: unknown): value is SpecificationSourceManifest =>
  validateSpecificationSourceManifest(value).length === 0;

const loadManifest = async (): Promise<SpecificationSourceManifest> => {
  const text = await Bun.file(MANIFEST_PATH).text();
  let value: unknown;
  try {
    value = JSON.parse(text);
  } catch (cause) {
    throw new SpecificationSourceError({
      message: "Specification manifest is not valid JSON",
      cause,
    });
  }
  const issues = validateSpecificationSourceManifest(value);
  if (issues.length > 0) {
    throw new SpecificationSourceError({
      message: `Specification manifest is invalid:\n${issues.map((issue) => `- ${issue}`).join("\n")}`,
    });
  }
  if (!isSpecificationSourceManifest(value)) {
    throw new SpecificationSourceError({ message: "Specification manifest validation failed" });
  }
  return value;
};

const resolveCachePath = (cachePath: string): string => {
  const resolved = path.resolve(REPOSITORY_ROOT, cachePath);
  if (!resolved.startsWith(`${CACHE_ROOT}${path.sep}`)) {
    throw new SpecificationSourceError({ message: `Cache path escapes cache root: ${cachePath}` });
  }
  return resolved;
};

const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

const readBoundedResponse = async (response: Response, maxBytes: number): Promise<Uint8Array> => {
  const contentLength = response.headers.get("content-length");
  if (contentLength !== null && Number(contentLength) > maxBytes) {
    throw new SpecificationSourceError({
      message: `Source response exceeds the ${maxBytes}-byte manifest limit`,
    });
  }
  if (response.body === null) {
    throw new SpecificationSourceError({ message: "Source response has no body" });
  }
  const chunks: Uint8Array[] = [];
  const reader = response.body.getReader();
  let totalBytes = 0;
  while (true) {
    // oxlint-disable-next-line no-await-in-loop -- each chunk is bounded before the next is read
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    totalBytes += value.length;
    if (totalBytes > maxBytes) {
      await reader.cancel();
      throw new SpecificationSourceError({
        message: `Source response exceeds the ${maxBytes}-byte manifest limit`,
      });
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return bytes;
};

const verifyArtifact = async (
  source: ArtifactSpecificationSource,
  fetchMissing: boolean,
): Promise<"missing" | "verified"> => {
  const cachePath = resolveCachePath(source.cachePath);
  if (!existsSync(cachePath) && fetchMissing) {
    const response = await fetch(source.url, {
      signal: AbortSignal.timeout(SOURCE_FETCH_TIMEOUT_MS),
    });
    if (!response.ok) {
      throw new SpecificationSourceError({
        message: `Failed to fetch ${source.id}: HTTP ${response.status}`,
      });
    }
    if (!response.url.startsWith("https://")) {
      throw new SpecificationSourceError({
        message: `Source redirected outside HTTPS: ${source.id}`,
      });
    }
    const bytes = await readBoundedResponse(response, source.bytes);
    if (bytes.length !== source.bytes || sha256(bytes) !== source.sha256) {
      throw new SpecificationSourceError({
        message: `Fetched artifact does not match manifest: ${source.id}`,
      });
    }
    await mkdir(path.dirname(cachePath), { recursive: true });
    await Bun.write(cachePath, bytes);
  }
  if (!existsSync(cachePath)) {
    return "missing";
  }
  const metadata = await stat(cachePath);
  if (!metadata.isFile() || metadata.size !== source.bytes) {
    throw new SpecificationSourceError({
      message: `Cached artifact does not match manifest: ${source.id}`,
    });
  }
  const bytes = new Uint8Array(await Bun.file(cachePath).arrayBuffer());
  if (bytes.length !== source.bytes || sha256(bytes) !== source.sha256) {
    throw new SpecificationSourceError({
      message: `Cached artifact does not match manifest: ${source.id}`,
    });
  }
  return "verified";
};

const runGit = async (args: string[], cwd = REPOSITORY_ROOT): Promise<string> => {
  const process = Bun.spawn(["git", ...args], {
    cwd,
    signal: AbortSignal.timeout(SOURCE_FETCH_TIMEOUT_MS),
    stdout: "pipe",
    stderr: "pipe",
  });
  const [exitCode, stdout, stderr] = await Promise.all([
    process.exited,
    new Response(process.stdout).text(),
    new Response(process.stderr).text(),
  ]);
  if (exitCode !== 0) {
    throw new SpecificationSourceError({
      message: `git ${args.at(0) ?? "command"} failed: ${stderr.trim()}`,
    });
  }
  return stdout.trim();
};

const verifyGitSource = async (
  source: GitSpecificationSource,
  fetchMissing: boolean,
): Promise<"missing" | "verified"> => {
  const cachePath = resolveCachePath(source.cachePath);
  const gitDirectory = path.join(cachePath, ".git");
  if (!existsSync(gitDirectory) && fetchMissing) {
    await rm(cachePath, { recursive: true, force: true });
    await mkdir(path.dirname(cachePath), { recursive: true });
    await runGit(["clone", "--filter=blob:none", "--no-checkout", source.repository, cachePath]);
  }
  if (!existsSync(gitDirectory)) {
    return "missing";
  }
  if (fetchMissing) {
    await runGit(["fetch", "--depth=1", "origin", source.commit], cachePath);
    await runGit(["checkout", "--detach", source.commit], cachePath);
  }
  const [origin, commit, tree] = await Promise.all([
    runGit(["remote", "get-url", "origin"], cachePath),
    runGit(["rev-parse", "HEAD"], cachePath),
    runGit(["rev-parse", "HEAD^{tree}"], cachePath),
  ]);
  if (origin !== source.repository || commit !== source.commit || tree !== source.tree) {
    throw new SpecificationSourceError({
      message: `Cached repository does not match manifest: ${source.id}`,
    });
  }
  return "verified";
};

const verifySource = async (
  source: SpecificationSource,
  fetchMissing: boolean,
): Promise<"missing" | "verified"> => {
  if (source.sourceType === "artifact") {
    return await verifyArtifact(source, fetchMissing);
  }
  return await verifyGitSource(source, fetchMissing);
};

const main = async (args: string[]): Promise<void> => {
  const command = args.at(0) ?? "check";
  if (command !== "check" && command !== "fetch") {
    throw new SpecificationSourceError({
      message: "Usage: bun scripts/specification-sources.ts [check|fetch]",
    });
  }
  const manifest = await loadManifest();
  const results = await Promise.all(
    manifest.sources.map(async (source) => ({
      id: source.id,
      status: await verifySource(source, command === "fetch"),
    })),
  );
  const verified = results.filter(({ status }) => status === "verified").length;
  const missing = results.length - verified;
  process.stdout.write(
    `Specification sources: ${manifest.sources.length} valid, ${verified} cached, ${missing} not cached\n`,
  );
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    const message = cause instanceof Error ? cause.message : String(cause);
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  });
}
