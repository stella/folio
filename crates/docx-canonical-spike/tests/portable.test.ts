import { expect, test } from "bun:test";
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { exportPortable, importPortable } from "../bench/portable";

test("portable import verifies the complete source and hash set before copying", () => {
  const directory = mkdtempSync(resolve(tmpdir(), "canonical-portable-"));
  try {
    const root = resolve(directory, "crates/spike");
    const artifact = resolve(directory, "artifact");
    for (const path of ["src", "bench", "tests", "../../packages/docx-core/src"])
      mkdirSync(resolve(root, path), { recursive: true });
    for (const name of ["Cargo.toml", "Cargo.lock", "tsconfig.json"])
      writeFileSync(resolve(root, name), name);
    for (const path of [
      "wasm-bindgen-release/docx_canonical_spike.js",
      "wasm-bindgen-release/docx_canonical_spike_bg.wasm",
      "wasm-bindgen-release/docx_canonical_spike.d.ts",
      "wasm-bindgen-release/docx_canonical_spike_bg.wasm.d.ts",
      "benchmark-browser/browser.js",
    ]) {
      const file = resolve(root, "target", path);
      mkdirSync(dirname(file), { recursive: true });
      writeFileSync(file, path);
    }
    exportPortable(root, artifact);
    const destination = resolve(root, "target/benchmark-browser/browser.js");
    writeFileSync(destination, "old");
    importPortable(root, artifact);
    expect(readFileSync(destination, "utf8")).toBe("benchmark-browser/browser.js");
    writeFileSync(destination, "keep");
    writeFileSync(
      resolve(artifact, "wasm-bindgen-release/docx_canonical_spike_bg.wasm"),
      "tampered",
    );
    expect(() => importPortable(root, artifact)).toThrow("hash/set mismatch");
    expect(readFileSync(destination, "utf8")).toBe("keep");
    exportPortable(root, artifact);
    writeFileSync(resolve(root, "src/changed.rs"), "changed source");
    expect(() => importPortable(root, artifact)).toThrow("does not match");
    expect(readFileSync(destination, "utf8")).toBe("keep");
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
