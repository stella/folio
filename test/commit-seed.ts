import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Seeds derived from the commit under test, so randomized tests search a new
 * sequence on every commit yet a rerun of the same commit replays exactly (a
 * red run stays red instead of passing on rerun). Used by the property tests
 * (test/property-testing.ts) and the consumer-scenario fuzz flows
 * (scripts/consumer-scenarios.ts) under CI.
 */

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

let cachedSha: string | null | undefined;

/** `GITHUB_SHA`, else `git rev-parse HEAD`, else `null`. */
export const commitSha = (): string | null => {
  if (cachedSha === undefined) {
    const fromEnv = process.env["GITHUB_SHA"];
    if (fromEnv !== undefined && fromEnv !== "") {
      cachedSha = fromEnv;
    } else {
      try {
        cachedSha = execFileSync("git", ["rev-parse", "HEAD"], {
          cwd: REPO_ROOT,
          encoding: "utf8",
          stdio: ["ignore", "pipe", "ignore"],
        }).trim();
      } catch {
        cachedSha = null;
      }
    }
  }
  return cachedSha;
};

/**
 * 32-bit FNV-1a over the UTF-16 code units, finished with murmur3's avalanche
 * so salts that differ in one character land far apart, as a signed integer.
 */
export const hash32 = (text: string): number => {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  hash ^= hash >>> 16;
  hash = Math.imul(hash, 0x85ebca6b);
  hash ^= hash >>> 13;
  hash = Math.imul(hash, 0xc2b2ae35);
  hash ^= hash >>> 16;
  return hash | 0;
};

/**
 * A signed 32-bit seed for `salt` at the current commit, or `undefined` when
 * the commit cannot be read.
 */
export const commitSeed = (salt: string): number | undefined => {
  const sha = commitSha();
  return sha === null ? undefined : hash32(`${sha}\0${salt}`);
};
