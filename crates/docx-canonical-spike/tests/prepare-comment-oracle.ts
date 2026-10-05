/** Materialize the pinned schema-10 TS oracle from a locally available commit. */
import { TaggedError } from "better-result";
import { mkdirSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { createHash } from "node:crypto";

export const COMMENT_ORACLE_COMMIT = "617c0f4703a6cf43adb7ecda34366c29b93cc60f";
export const COMMENT_ORACLE_SOURCE = "packages/docx-core/src";
const repository = fileURLToPath(new URL("../../../", import.meta.url));
const directory = new URL(`../target/comment-oracle/${COMMENT_ORACLE_COMMIT}/`, import.meta.url);
const manifest = new URL("provenance.json", directory);

class CommentOraclePreparationError extends TaggedError("CommentOraclePreparationError")<{
  message: string;
}> {}

export const commentOracleModule = (path: string) =>
  new URL(`${COMMENT_ORACLE_SOURCE}/${path}`, directory);

export const assertCommentOraclePrepared = (): void => {
  if (!existsSync(manifest))
    throw new CommentOraclePreparationError({
      message: `Prepare the pinned comment oracle first: bun ${fileURLToPath(import.meta.url)}. The script never fetches source.`,
    });
  // Boundary parsing: malformed provenance must fail before importing an oracle.
  const expected = JSON.stringify({
    commit: COMMENT_ORACLE_COMMIT,
    source: COMMENT_ORACLE_SOURCE,
    schema: 10,
  });
  if (readFileSync(manifest, "utf8").trim() !== expected)
    throw new CommentOraclePreparationError({
      message: "Comment oracle provenance does not match the pinned commit and schema.",
    });
  for (const file of [
    "ops/apply.ts",
    "ops/comments.ts",
    "ops/types.ts",
    "ops/wire.ts",
    "model/document.ts",
  ])
    if (!existsSync(commentOracleModule(file)))
      throw new CommentOraclePreparationError({
        message: `The pinned comment oracle is incomplete: ${file}.`,
      });
  const tree = spawnSync("git", ["ls-tree", "-rz", COMMENT_ORACLE_COMMIT, COMMENT_ORACLE_SOURCE], {
    cwd: repository,
    encoding: "utf8",
    maxBuffer: 8 * 1024 * 1024,
  });
  if (tree.error || tree.status !== 0)
    throw new CommentOraclePreparationError({
      message: "The exact comment oracle commit must be present locally to verify provenance.",
    });
  for (const entry of tree.stdout.split("\0").filter(Boolean)) {
    const match = /^100(?:644|755) blob ([0-9a-f]{40})\t(.+)$/u.exec(entry);
    const hash = match?.at(1);
    const path = match?.at(2);
    if (!hash || !path || !path.startsWith(`${COMMENT_ORACLE_SOURCE}/`))
      throw new CommentOraclePreparationError({
        message: "Unexpected source entry in the pinned comment oracle.",
      });
    const file = new URL(path, directory);
    if (!existsSync(file))
      throw new CommentOraclePreparationError({
        message: `The pinned comment oracle is incomplete: ${path}.`,
      });
    const bytes = readFileSync(file);
    const actual = createHash("sha1").update(`blob ${bytes.length}\0`).update(bytes).digest("hex");
    if (actual !== hash)
      throw new CommentOraclePreparationError({
        message: `Pinned comment source differs from ${COMMENT_ORACLE_COMMIT}: ${path}.`,
      });
  }
};

export const prepareCommentOracle = (): void => {
  const available = spawnSync(
    "git",
    ["cat-file", "-e", `${COMMENT_ORACLE_COMMIT}:${COMMENT_ORACLE_SOURCE}/ops/comments.ts`],
    { cwd: repository, encoding: "utf8" },
  );
  if (available.error || available.status !== 0)
    throw new CommentOraclePreparationError({
      message: `Pinned comment source ${COMMENT_ORACLE_COMMIT} is absent locally; fetch that exact commit before preparing the oracle. No network request was made.`,
    });
  const archive = spawnSync(
    "git",
    ["archive", "--format=tar", COMMENT_ORACLE_COMMIT, COMMENT_ORACLE_SOURCE],
    { cwd: repository, maxBuffer: 64 * 1024 * 1024 },
  );
  if (archive.error || archive.status !== 0)
    throw new CommentOraclePreparationError({
      message: `Cannot archive pinned comment source: ${archive.error?.message ?? archive.stderr.toString()}`,
    });
  // The only removed tree is this generated, commit-named oracle directory.
  rmSync(directory, { recursive: true, force: true });
  mkdirSync(directory, { recursive: true });
  const extracted = spawnSync("tar", ["-x", "-C", fileURLToPath(directory)], {
    input: archive.stdout,
    encoding: "utf8",
  });
  if (extracted.error || extracted.status !== 0)
    throw new CommentOraclePreparationError({
      message: `Cannot extract pinned comment source: ${extracted.error?.message ?? extracted.stderr}`,
    });
  writeFileSync(
    manifest,
    `${JSON.stringify({ commit: COMMENT_ORACLE_COMMIT, source: COMMENT_ORACLE_SOURCE, schema: 10 })}\n`,
  );
  assertCommentOraclePrepared();
};

if (import.meta.main) {
  prepareCommentOracle();
  process.stdout.write(`Prepared comment oracle ${COMMENT_ORACLE_COMMIT} (schema 10).\n`);
}
