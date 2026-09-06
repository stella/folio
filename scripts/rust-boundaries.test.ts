import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import ts from "typescript";

const REPOSITORY_ROOT = path.join(import.meta.dir, "..");

/**
 * Every artifact compiled from Rust, and the module that owns the boundary to
 * it. The lint rule keeps a second entry point from appearing; this keeps the
 * boundary module itself thin, which is the other half of the same contract:
 * what the crate does has one implementation, and the TypeScript beside it
 * initializes the artifact and translates its errors.
 */
const RUST_BOUNDARIES = [
  {
    label: "DOCX projection",
    module: "packages/docx-core/src/projection.ts",
    imports: ["better-result", "./generated/docx_kernel.js"],
    // A second OOXML or archive reader here would be a second answer to the
    // question the kernel exists to answer.
    forbidden: ["jszip", "fast-xml-parser", "DOMParser"],
  },
  {
    label: "text shaping",
    module: "packages/core/src/shaping/shaper.ts",
    imports: ["better-result", "../generated/text_shaper.js"],
    // Shaping in TypeScript beside the boundary is how the measurer and the
    // painter would come to disagree about which glyphs a word is.
    forbidden: ["harfbuzz", "opentype", "fontkit"],
  },
] as const;

describe("Rust boundary modules stay thin", () => {
  for (const { label, module, imports, forbidden } of RUST_BOUNDARIES) {
    test(`the ${label} boundary imports only its artifact`, async () => {
      const source = await readFile(path.join(REPOSITORY_ROOT, module), "utf8");
      const importedModules = ts
        .preProcessFile(source)
        .importedFiles.map(({ fileName }) => fileName);

      expect(importedModules).toEqual([...imports]);
      for (const name of forbidden) {
        expect(source).not.toContain(name);
      }
    });
  }
});
