/**
 * Fetch, verify and lock the public DOCX corpus.
 *
 * Each source is a public repository pinned to a commit, taken through a sparse
 * partial clone so only the `.docx` paths the manifest names are ever
 * downloaded. Content lands in a cache directory outside the repository and is
 * never committed; `corpus/sources.lock.json` records a relative path, a
 * SHA-256 and a byte count per file, which is what makes a run reproducible and
 * lets CI key its cache on the exact corpus.
 *
 * Usage:
 *   bun scripts/corpus-sources.ts fetch   # download into the cache, then relock
 *   bun scripts/corpus-sources.ts lock    # rewrite the lock from the cache
 *   bun scripts/corpus-sources.ts check   # verify the cache against the lock
 */

import { existsSync } from "node:fs";
import { mkdir, stat } from "node:fs/promises";
import path from "node:path";

import {
  CorpusManifestError,
  type CorpusLockEntry,
  type CorpusLockSource,
  type CorpusManifest,
  type CorpusSource,
  type CorpusTier,
  LOCK_PATH,
  corpusLockDigest,
  loadCorpusLock,
  loadCorpusManifest,
  sha256Bytes,
  sourceCheckoutPath,
  writeJsonFile,
} from "./lib/corpus-manifest";

const GIT_TIMEOUT_MS = 300_000;
const FETCH_CONCURRENCY = 3;

const runGit = async (args: string[], cwd: string): Promise<string> => {
  const child = Bun.spawn(["git", ...args], {
    cwd,
    signal: AbortSignal.timeout(GIT_TIMEOUT_MS),
    stdout: "pipe",
    stderr: "pipe",
  });
  const [exitCode, stdout, stderr] = await Promise.all([
    child.exited,
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
  ]);
  if (exitCode !== 0) {
    throw new CorpusManifestError({
      message: `git ${args.at(0) ?? "command"} failed in ${cwd}: ${stderr.trim()}`,
    });
  }
  return stdout.trim();
};

/**
 * Materialise one source at its pinned commit.
 *
 * `--filter=blob:none` with a sparse-checkout pattern set before the checkout
 * means git downloads the commit's trees and then only the blobs the patterns
 * match: a few megabytes of `.docx` out of repositories that are gigabytes
 * whole.
 */
const fetchSource = async (source: CorpusSource): Promise<void> => {
  const checkout = sourceCheckoutPath(source.id);
  await mkdir(checkout, { recursive: true });
  if (!existsSync(path.join(checkout, ".git"))) {
    await runGit(["init", "--quiet"], checkout);
    await runGit(["remote", "add", "origin", source.repository], checkout);
  }
  const origin = await runGit(["remote", "get-url", "origin"], checkout);
  if (origin !== source.repository) {
    throw new CorpusManifestError({
      message: `Cached checkout for ${source.id} points at ${origin}, not ${source.repository}`,
    });
  }
  await runGit(["sparse-checkout", "set", "--no-cone", ...source.paths], checkout);
  await runGit(["fetch", "--depth=1", "--filter=blob:none", "origin", source.commit], checkout);
  await runGit(["checkout", "--detach", "--force", source.commit], checkout);
  await verifySourceCheckout(source);
};

const verifySourceCheckout = async (source: CorpusSource): Promise<void> => {
  const checkout = sourceCheckoutPath(source.id);
  const [commit, tree] = await Promise.all([
    runGit(["rev-parse", "HEAD"], checkout),
    runGit(["rev-parse", "HEAD^{tree}"], checkout),
  ]);
  if (commit !== source.commit || tree !== source.tree) {
    throw new CorpusManifestError({
      message: `Cached checkout for ${source.id} is at ${commit} (${tree}), not the pinned ${source.commit} (${source.tree})`,
    });
  }
};

const mapWithConcurrency = async <T, R>(
  items: readonly T[],
  limit: number,
  run: (item: T) => Promise<R>,
): Promise<R[]> => {
  const results: R[] = Array.from({ length: items.length });
  let next = 0;
  const worker = async (): Promise<void> => {
    while (next < items.length) {
      const index = next;
      next += 1;
      // oxlint-disable-next-line no-await-in-loop -- a worker is sequential by design; `limit` of them run in parallel
      results[index] = await run(items[index] as T);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
};

type SourceFiles = {
  id: string;
  commit: string;
  tier: CorpusTier;
  files: CorpusLockEntry[];
  oversized: number;
};

const collectSourceFiles = async (
  source: CorpusSource,
  maxFileBytes: number,
): Promise<SourceFiles> => {
  const checkout = sourceCheckoutPath(source.id);
  const relativePaths: string[] = [];
  for await (const relative of new Bun.Glob("**/*.docx").scan({ cwd: checkout, onlyFiles: true })) {
    relativePaths.push(relative);
  }
  relativePaths.sort();

  const files: CorpusLockEntry[] = [];
  let oversized = 0;
  for (const relative of relativePaths) {
    // oxlint-disable-next-line no-await-in-loop -- hashing is bounded by file size, and a whole source is one bounded set
    const bytes = new Uint8Array(await Bun.file(path.join(checkout, relative)).arrayBuffer());
    if (bytes.byteLength > maxFileBytes) {
      oversized += 1;
      continue;
    }
    files.push({ path: relative, sha256: sha256Bytes(bytes), bytes: bytes.byteLength });
  }
  return { id: source.id, commit: source.commit, tier: source.tier, files, oversized };
};

type BuiltLock = {
  lock: {
    schemaVersion: 1;
    manifestDigest: string;
    fileCount: number;
    totalBytes: number;
    sources: CorpusLockSource[];
  };
  oversized: number;
  dropped: number;
};

const buildLock = async (manifest: CorpusManifest, digest: string): Promise<BuiltLock> => {
  const collected = await mapWithConcurrency(manifest.sources, FETCH_CONCURRENCY, (source) =>
    collectSourceFiles(source, manifest.maxFileBytes),
  );

  // The file cap is applied by ascending tier, then in manifest order, so the
  // corpus a lock describes is a function of the manifest alone and a large
  // tier-2 source can never displace the tier-1 files CI runs.
  let remaining = manifest.fileLimit;
  let dropped = 0;
  const keptById = new Map<string, CorpusLockEntry[]>();
  for (const { id, files } of [...collected].sort((left, right) => left.tier - right.tier)) {
    const kept = files.slice(0, Math.max(remaining, 0));
    dropped += files.length - kept.length;
    remaining -= kept.length;
    keptById.set(id, kept);
  }

  const sources: CorpusLockSource[] = [];
  let fileCount = 0;
  let totalBytes = 0;
  for (const { id, commit, tier } of collected) {
    const kept = keptById.get(id) ?? [];
    fileCount += kept.length;
    for (const file of kept) {
      totalBytes += file.bytes;
    }
    sources.push({ id, commit, tier, files: kept });
  }

  return {
    lock: { schemaVersion: 1, manifestDigest: digest, fileCount, totalBytes, sources },
    oversized: collected.reduce((total, source) => total + source.oversized, 0),
    dropped,
  };
};

const writeLock = async (manifest: CorpusManifest, digest: string): Promise<void> => {
  const { lock, oversized, dropped } = await buildLock(manifest, digest);
  await writeJsonFile(LOCK_PATH, lock);
  const megabytes = (lock.totalBytes / 1_048_576).toFixed(1);
  process.stdout.write(
    `Corpus lock: ${lock.fileCount} files, ${megabytes} MB, digest ${corpusLockDigest(lock).slice(0, 12)}` +
      `${oversized > 0 ? `, ${oversized} over the size limit` : ""}` +
      `${dropped > 0 ? `, ${dropped} over the file limit` : ""}\n`,
  );
};

type CheckCounts = { verified: number; missing: number; mismatched: string[] };

const checkLock = async (): Promise<CheckCounts> => {
  const [{ manifest, digest }, lock] = await Promise.all([loadCorpusManifest(), loadCorpusLock()]);
  if (lock.manifestDigest !== digest) {
    throw new CorpusManifestError({
      message:
        "corpus/sources.lock.json was built from a different manifest. Run `bun run corpus:fetch`.",
    });
  }
  const lockedIds = new Set(lock.sources.map((source) => source.id));
  const manifestIds = manifest.sources.map((source) => source.id);
  const unlocked = manifestIds.filter((id) => !lockedIds.has(id));
  if (unlocked.length > 0) {
    throw new CorpusManifestError({
      message: `Sources missing from the lock: ${unlocked.join(", ")}`,
    });
  }

  const counts: CheckCounts = { verified: 0, missing: 0, mismatched: [] };
  for (const source of lock.sources) {
    const checkout = sourceCheckoutPath(source.id);
    for (const file of source.files) {
      const absolute = path.join(checkout, file.path);
      if (!existsSync(absolute)) {
        counts.missing += 1;
        continue;
      }
      // oxlint-disable-next-line no-await-in-loop -- one bounded read per locked file
      const metadata = await stat(absolute);
      if (metadata.size !== file.bytes) {
        counts.mismatched.push(
          `${source.id}/${file.path}: size ${metadata.size}, locked ${file.bytes}`,
        );
        continue;
      }
      // oxlint-disable-next-line no-await-in-loop -- see above
      const bytes = new Uint8Array(await Bun.file(absolute).arrayBuffer());
      if (sha256Bytes(bytes) !== file.sha256) {
        counts.mismatched.push(`${source.id}/${file.path}: content differs from the lock`);
        continue;
      }
      counts.verified += 1;
    }
  }
  return counts;
};

const main = async (args: string[]): Promise<void> => {
  const command = args.at(0) ?? "check";
  if (command !== "check" && command !== "fetch" && command !== "lock") {
    throw new CorpusManifestError({
      message: "Usage: bun scripts/corpus-sources.ts [check|fetch|lock]",
    });
  }

  if (command === "check") {
    const { verified, missing, mismatched } = await checkLock();
    if (mismatched.length > 0) {
      throw new CorpusManifestError({
        message: `Cached corpus does not match the lock:\n${mismatched
          .slice(0, 20)
          .map((issue) => `- ${issue}`)
          .join("\n")}`,
      });
    }
    process.stdout.write(`Corpus cache: ${verified} verified, ${missing} not cached\n`);
    return;
  }

  const { manifest, digest } = await loadCorpusManifest();
  if (command === "fetch") {
    await mapWithConcurrency(manifest.sources, FETCH_CONCURRENCY, fetchSource);
  } else {
    await Promise.all(manifest.sources.map(verifySourceCheckout));
  }
  await writeLock(manifest, digest);
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    process.stderr.write(`${cause instanceof Error ? cause.message : String(cause)}\n`);
    process.exitCode = 1;
  });
}
