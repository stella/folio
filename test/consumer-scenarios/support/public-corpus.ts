/** External SHA-addressed fixtures; third-party package bytes stay in the cache. */
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { TaggedError } from "better-result";

export const PUBLIC_CORPUS_PREFIX = "public-corpus:";
export type PublicCorpusFixture = `${typeof PUBLIC_CORPUS_PREFIX}${string}`;
export const PUBLIC_CORPUS_DIRECTORY_ENV = "FOLIO_SCENARIO_PUBLIC_CORPUS_DIR";
const SHA256 = /^[a-f0-9]{64}$/u;
export const MAX_PUBLIC_CORPUS_BYTES = 128 * 1024;
export const MAX_PUBLIC_CORPUS_FILES = 8;

class PublicCorpusFixtureError extends TaggedError("PublicCorpusFixtureError")<{
  message: string;
}> {}

export const isPublicCorpusFixture = (fixture: string): fixture is PublicCorpusFixture =>
  fixture.startsWith(PUBLIC_CORPUS_PREFIX);

const fixtureDirectory = (): string => {
  const directory = process.env[PUBLIC_CORPUS_DIRECTORY_ENV];
  if (!directory || !path.isAbsolute(directory)) {
    throw new PublicCorpusFixtureError({
      message: `${PUBLIC_CORPUS_DIRECTORY_ENV} must name the absolute staged corpus cache; run bun run corpus:fetch and bun scripts/stage-public-corpus-flows.ts first`,
    });
  }
  return directory;
};

export const loadPublicCorpusFixture = async (
  fixture: PublicCorpusFixture,
): Promise<Uint8Array> => {
  const sha256 = fixture.slice(PUBLIC_CORPUS_PREFIX.length);
  if (!SHA256.test(sha256)) {
    throw new PublicCorpusFixtureError({ message: "Public corpus fixture requires a SHA-256" });
  }
  const bytes = new Uint8Array(await readFile(path.join(fixtureDirectory(), `${sha256}.docx`)));
  if (
    bytes.byteLength > MAX_PUBLIC_CORPUS_BYTES ||
    createHash("sha256").update(bytes).digest("hex") !== sha256
  ) {
    throw new PublicCorpusFixtureError({
      message: `Public corpus fixture ${sha256} differs from its pinned bytes or exceeds the size limit`,
    });
  }
  return bytes;
};

export const publicCorpusFixtures = async (): Promise<PublicCorpusFixture[]> => {
  const value: unknown = JSON.parse(
    await readFile(path.join(fixtureDirectory(), "index.json"), "utf8"),
  );
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_PUBLIC_CORPUS_FILES) {
    throw new PublicCorpusFixtureError({
      message: "Public corpus index requires 1–8 hashed fixtures",
    });
  }
  const hashes: unknown[] = value;
  const fixtures: PublicCorpusFixture[] = [];
  for (const hash of hashes) {
    if (typeof hash !== "string" || !SHA256.test(hash)) {
      throw new PublicCorpusFixtureError({
        message: "Public corpus index contains an invalid SHA-256",
      });
    }
    fixtures.push(`${PUBLIC_CORPUS_PREFIX}${hash}`);
  }
  if (new Set(fixtures).size !== fixtures.length) {
    throw new PublicCorpusFixtureError({ message: "Public corpus index repeats a fixture" });
  }
  return fixtures;
};
