#!/usr/bin/env bun

import {
  buildModeFrom,
  buildRustWasmArtifact,
  type RustWasmArtifact,
} from "./lib/rust-wasm-artifact.ts";

const mode = buildModeFrom(process.argv[2], "bun scripts/build-docx-kernel-wasm.ts");

// Review content and UTF-8/UTF-16 spans have an explicit additive budget so
// unrelated feature growth cannot consume their allowance unnoticed.
const maximumReviewDetailBytes = 11 * 1024;
const maximumReviewDetailBrotliBytes = 2 * 1024;
// Direct paragraph styles and resolved outline levels have their own bounded
// allowance so future projection growth cannot consume it silently.
const maximumParagraphStructureBytes = 2 * 1024;
const maximumParagraphStructureBrotliBytes = 1024;
// Effective run formatting, style-cascade resolution, and visible non-text run
// fallbacks have an explicit allowance so future projection growth stays bounded.
const maximumEffectiveTextBytes = 11 * 1024;
const maximumEffectiveTextBrotliBytes = 5 * 1024;
// Unicode-script classification lets mixed-script runs select b versus bCs at
// exact UTF-16 boundaries without host calls or a second document pass.
const maximumComplexScriptFormattingBytes = 29 * 1024;
const maximumComplexScriptFormattingBrotliBytes = 8 * 1024;
// Prepared paragraph-style results prevent every paragraph and run from walking
// the same style cascade again during projection and structural materialization.
const maximumPreparedStylesBytes = 3 * 1024;
const maximumPreparedStylesBrotliBytes = 1024;
// Paragraph-boundary bookmark resolution has its own allowance so accepting
// schema-valid markers outside paragraphs cannot hide unrelated parser growth.
const maximumBookmarkBoundaryBytes = 4 * 1024;
const maximumBookmarkBoundaryBrotliBytes = 2 * 1024;
// Bounded numbering parsing and effective indentation resolution have a
// separate allowance so later projection work cannot silently consume it.
const maximumNumberingStructureBytes = 13 * 1024;
const maximumNumberingStructureBrotliBytes = 5 * 1024;
// Effective `w:jc` resolution across the paragraph, its style chain, and
// document defaults has its own allowance so later projection work cannot
// silently consume it.
const maximumParagraphAlignmentBytes = 2 * 1024;
const maximumParagraphAlignmentBrotliBytes = 1024;

const kernel = {
  label: "DOCX kernel",
  crate: "stella-docx-kernel",
  crateDirectory: "crates/docx-kernel",
  cargoArtifact: "stella_docx_kernel",
  outName: "docx_kernel",
  generatedDirectory: "packages/docx-core/src/generated",
  regenerateCommand: "bun --filter @stll/docx-core wasm:generate",
  buildScript: "scripts/build-docx-kernel-wasm.ts",
  maximumWasmBytes:
    222 * 1024 +
    maximumReviewDetailBytes +
    maximumParagraphStructureBytes +
    maximumEffectiveTextBytes +
    maximumComplexScriptFormattingBytes +
    maximumPreparedStylesBytes +
    maximumBookmarkBoundaryBytes +
    maximumNumberingStructureBytes +
    maximumParagraphAlignmentBytes,
  maximumBrotliBytes:
    100 * 1024 +
    maximumReviewDetailBrotliBytes +
    maximumParagraphStructureBrotliBytes +
    maximumEffectiveTextBrotliBytes +
    maximumComplexScriptFormattingBrotliBytes +
    maximumPreparedStylesBrotliBytes +
    maximumBookmarkBoundaryBrotliBytes +
    maximumNumberingStructureBrotliBytes +
    maximumParagraphAlignmentBrotliBytes,
} as const satisfies RustWasmArtifact;

await buildRustWasmArtifact(kernel, mode);
