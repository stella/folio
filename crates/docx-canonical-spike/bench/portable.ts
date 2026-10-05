/** Portable build output carries relative file names, source identity and hashes. */
import { createHash } from "node:crypto";
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { sourceDigest } from "./artifacts";

const PORTABLE_FILES = [
  "wasm-bindgen-release/docx_canonical_spike.js",
  "wasm-bindgen-release/docx_canonical_spike_bg.wasm",
  "wasm-bindgen-release/docx_canonical_spike.d.ts",
  "wasm-bindgen-release/docx_canonical_spike_bg.wasm.d.ts",
  "benchmark-browser/browser.js",
] as const;
const hashFile = (path: string) => createHash("sha256").update(readFileSync(path)).digest("hex");

export const exportPortable = (root: string, directory: string) => {
  const files = PORTABLE_FILES.map((path) => {
    const source = resolve(root, "target", path);
    const destination = resolve(directory, path);
    mkdirSync(dirname(destination), { recursive: true });
    copyFileSync(source, destination);
    return { path, sha256: hashFile(destination) };
  });
  writeFileSync(
    resolve(directory, "portable-manifest.json"),
    `${JSON.stringify({ status: "built", sourceHash: sourceDigest(root), files }, null, 2)}\n`,
  );
};

export const importPortable = (root: string, directory: string) => {
  const manifest: unknown = JSON.parse(
    readFileSync(resolve(directory, "portable-manifest.json"), "utf8"),
  );
  if (
    typeof manifest !== "object" ||
    manifest === null ||
    !("status" in manifest) ||
    manifest.status !== "built" ||
    !("sourceHash" in manifest) ||
    manifest.sourceHash !== sourceDigest(root) ||
    !("files" in manifest) ||
    !Array.isArray(manifest.files) ||
    manifest.files.length !== PORTABLE_FILES.length
  )
    throw new TypeError("Portable WASM does not match this source tree.");
  // Verify the complete set before copying any file; paths come from this owner table.
  for (const path of PORTABLE_FILES) {
    const matches = manifest.files.filter(
      (file: unknown) =>
        typeof file === "object" && file !== null && "path" in file && file.path === path,
    );
    const entry: unknown = matches.at(0);
    if (
      matches.length !== 1 ||
      typeof entry !== "object" ||
      entry === null ||
      !("sha256" in entry) ||
      typeof entry.sha256 !== "string" ||
      hashFile(resolve(directory, path)) !== entry.sha256
    )
      throw new TypeError(`Portable artifact hash/set mismatch: ${path}.`);
  }
  for (const path of PORTABLE_FILES) {
    const destination = resolve(root, "target", path);
    mkdirSync(dirname(destination), { recursive: true });
    copyFileSync(resolve(directory, path), destination);
  }
  return manifest;
};
