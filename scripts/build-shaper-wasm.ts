#!/usr/bin/env bun

import {
  buildModeFrom,
  buildRustWasmArtifact,
  type RustWasmArtifact,
} from "./lib/rust-wasm-artifact.ts";

const mode = buildModeFrom(process.argv[2], "bun scripts/build-shaper-wasm.ts");

// The shaper is its own artifact with its own budget rather than part of the
// DOCX kernel, and it is fetched the first time a run needs shaping. A document
// in Latin, Cyrillic or Greek never loads it, so this size is not paid by every
// consumer the way the kernel's is.
//
// Most of it is the OpenType layout engine and the Unicode data it needs:
// script and joining properties, the Arabic and Indic shapers, and the
// normalizer. That work has no smaller correct form; a table-driven substitute
// would be the same data with the lookups reimplemented.
const maximumWasmBytes = 552 * 1024;
const maximumBrotliBytes = 198 * 1024;

const shaper = {
  label: "text shaper",
  crate: "stella-text-shaper",
  crateDirectory: "crates/text-shaper",
  cargoArtifact: "stella_text_shaper",
  outName: "text_shaper",
  generatedDirectory: "packages/core/src/generated",
  regenerateCommand: "bun --filter @stll/folio-core wasm:generate",
  buildScript: "scripts/build-shaper-wasm.ts",
  maximumWasmBytes,
  maximumBrotliBytes,
} as const satisfies RustWasmArtifact;

await buildRustWasmArtifact(shaper, mode);
