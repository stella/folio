/** Bind measurements to actual source bytes as well as build output bytes. */
import { createHash } from "node:crypto";
import { readdirSync, readFileSync } from "node:fs";
import { join, resolve, relative } from "node:path";

export const sourceDigest = (crateRoot: string) => {
  const repoRoot = resolve(crateRoot, "../..");
  const paths: string[] = [];
  const visit = (directory: string) => {
    for (const entry of readdirSync(directory, { withFileTypes: true })) {
      const path = join(directory, entry.name);
      if (entry.isDirectory()) visit(path);
      else if (entry.isFile()) paths.push(path);
    }
  };
  for (const directory of ["src", "bench", "tests"]) visit(resolve(crateRoot, directory));
  visit(resolve(repoRoot, "packages/docx-core/src"));
  paths.push(
    resolve(crateRoot, "Cargo.toml"),
    resolve(crateRoot, "Cargo.lock"),
    resolve(crateRoot, "tsconfig.json"),
  );
  const hash = createHash("sha256");
  for (const path of paths.sort())
    hash.update(relative(repoRoot, path)).update("\0").update(readFileSync(path)).update("\0");
  return hash.digest("hex");
};
