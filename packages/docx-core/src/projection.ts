import { TaggedError } from "better-result";

import initializeRuntime, {
  projectCompressedDocx as projectCompressedDocxInWasm,
  projectCompressedDocxWithReviewFacts as projectCompressedDocxWithReviewFactsInWasm,
  type DocxAttributedComment,
  type DocxAttributedRevision,
  type DocxCommentContent,
  type DocxPackageProjectionWire,
  type DocxProjectionBookmarkFact,
  type DocxProjectionFactSet,
  type DocxProjectionFormattingSpan,
  type DocxProjectionIndentationFact,
  type DocxProjectionNumberingFact,
  type DocxProjectionParagraph,
  type DocxProjectionReferenceFact,
  type DocxProjectionRevisionStatus,
  type DocxProjectionRevisionUnsupportedReason,
  type DocxProjectionStructuralFacts,
  type DocxProjectionStructuralSpan,
  type DocxProjectionStructure,
  type DocxProjectionUnknownReason,
  type DocxProjectionWire,
  type DocxReviewFactsWire,
  type DocxReviewFactSet,
  type DocxReviewDetail,
  type DocxReviewPoint,
  type DocxReviewSpan,
  type DocxReviewUnknownReason,
  type DocxRevisionContent,
} from "./generated/docx_kernel.js";

export type {
  DocxAttributedComment,
  DocxAttributedRevision,
  DocxCommentContent,
  DocxPackageProjectionWire,
  DocxProjectionBookmarkFact,
  DocxProjectionFactSet,
  DocxProjectionFormattingSpan,
  DocxProjectionIndentationFact,
  DocxProjectionNumberingFact,
  DocxProjectionParagraph,
  DocxProjectionReferenceFact,
  DocxProjectionRevisionStatus,
  DocxProjectionRevisionUnsupportedReason,
  DocxProjectionStructuralFacts,
  DocxProjectionStructuralSpan,
  DocxProjectionStructure,
  DocxProjectionUnknownReason,
  DocxProjectionWire,
  DocxReviewFactsWire,
  DocxReviewFactSet,
  DocxReviewDetail,
  DocxReviewPoint,
  DocxReviewSpan,
  DocxReviewUnknownReason,
  DocxRevisionContent,
};

export class DocxProjectionInitializationError extends TaggedError(
  "DocxProjectionInitializationError",
)<{
  message: string;
  cause: unknown;
}>() {}

export class DocxProjectionError extends TaggedError("DocxProjectionError")<{
  message: string;
  cause: unknown;
}>() {}

export type DocxProjectionWasmSource =
  | RequestInfo
  | URL
  | Response
  | BufferSource
  | WebAssembly.Module;

export type InitializeDocxProjectionOptions = {
  /**
   * Optional explicit WebAssembly source for runtimes that do not load assets
   * by URL. Browsers normally omit this and use the package-relative asset.
   */
  wasm?: DocxProjectionWasmSource;
};

let initialization: Promise<void> | undefined;

/** Loads and caches the single-threaded WebAssembly runtime. */
export const initializeDocxProjection = ({
  wasm,
}: InitializeDocxProjectionOptions = {}): Promise<void> => {
  initialization ??= initializeRuntime(wasm === undefined ? undefined : { module_or_path: wasm })
    .then(() => undefined)
    .catch((cause: unknown) => {
      initialization = undefined;
      throw new DocxProjectionInitializationError({
        message: "Could not initialize the DOCX projection runtime",
        cause,
      });
    });
  return initialization;
};

/**
 * Projects compressed DOCX bytes through the canonical Rust implementation.
 * The TypeScript boundary initializes WebAssembly and preserves its versioned
 * result; it does not contain an alternate OOXML parser.
 */
export const projectCompressedDocx = async (bytes: Uint8Array): Promise<DocxProjectionWire> => {
  await initializeDocxProjection();
  try {
    return projectCompressedDocxInWasm(bytes);
  } catch (cause) {
    throw new DocxProjectionError({
      message: "Could not project the DOCX package",
      cause,
    });
  }
};

/**
 * Projects the document snapshot and attributed review facts through one
 * bounded package scan. Optional review-part failures remain explicit unknown
 * fact families rather than invalidating the document snapshot.
 */
export const projectCompressedDocxWithReviewFacts = async (
  bytes: Uint8Array,
): Promise<DocxPackageProjectionWire> => {
  await initializeDocxProjection();
  try {
    return projectCompressedDocxWithReviewFactsInWasm(bytes);
  } catch (cause) {
    throw new DocxProjectionError({
      message: "Could not project the DOCX package",
      cause,
    });
  }
};
