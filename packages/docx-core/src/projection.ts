import { TaggedError } from "better-result";

import initializeRuntime, {
  projectCompressedDocx as projectCompressedDocxInWasm,
  projectCompressedDocxWithReadableReviewFacts as projectCompressedDocxWithReadableReviewFactsInWasm,
  projectCompressedDocxWithReviewFacts as projectCompressedDocxWithReviewFactsInWasm,
  type DocxAttributedComment,
  type DocxAttributedRevision,
  type DocxPackageProjectionWire,
  type DocxProjectionAlignment,
  type DocxProjectionAlignmentSource,
  type DocxProjectionAlignmentValue,
  type DocxProjectionBookmarkFact,
  type DocxProjectionFactSet,
  type DocxProjectionFormattingSpan,
  type DocxProjectionFormattingStatus,
  type DocxProjectionFormattingUnknownReason,
  type DocxProjectionIndentationFact,
  type DocxProjectionNumberingFact,
  type DocxProjectionOutlineLevelFact,
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
  type DocxReviewUnknownReason,
  type DocxRevisionKind,
} from "./generated/docx_kernel.js";

export type {
  DocxAttributedComment,
  DocxAttributedRevision,
  DocxPackageProjectionWire,
  DocxProjectionAlignment,
  DocxProjectionAlignmentSource,
  DocxProjectionAlignmentValue,
  DocxProjectionBookmarkFact,
  DocxProjectionFactSet,
  DocxProjectionFormattingSpan,
  DocxProjectionFormattingStatus,
  DocxProjectionFormattingUnknownReason,
  DocxProjectionIndentationFact,
  DocxProjectionNumberingFact,
  DocxProjectionOutlineLevelFact,
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
  DocxReviewUnknownReason,
  DocxRevisionKind,
};

export class DocxProjectionInitializationError extends TaggedError(
  "DocxProjectionInitializationError",
)<{
  message: string;
  cause: unknown;
}> {}

export class DocxProjectionError extends TaggedError("DocxProjectionError")<{
  message: string;
  cause: unknown;
}> {}

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

export type ProjectCompressedDocxOptions = {
  /** Defaults to resolved style highlighting; direct selects run markup only. */
  highlightProjection?: "direct" | "resolved";
};

/** Controls formatting and text materialization for fused package projection. */
export type ProjectCompressedDocxWithReviewFactsOptions = ProjectCompressedDocxOptions & {
  /** Selects host-coordinate controls or normalized readable text. */
  textMaterialization?: "word-host" | "readable-plain-text";
};

let initialization: Promise<void> | undefined;
const DOCUMENT_PROJECTION_FAILURE_MESSAGE = "Could not project the DOCX document";
const PACKAGE_PROJECTION_FAILURE_MESSAGE = "Could not project the DOCX package with review facts";

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
const projectWith = async <T>({
  bytes,
  project,
  message,
}: {
  bytes: Uint8Array;
  project: (input: Uint8Array) => T;
  message: string;
}): Promise<T> => {
  await initializeDocxProjection();
  try {
    return project(bytes);
  } catch (cause) {
    throw new DocxProjectionError({ message, cause });
  }
};

export const projectCompressedDocx = (
  bytes: Uint8Array,
  { highlightProjection = "resolved" }: ProjectCompressedDocxOptions = {},
): Promise<DocxProjectionWire> =>
  projectWith({
    bytes,
    project: (input) => projectCompressedDocxInWasm(input, highlightProjection),
    message: DOCUMENT_PROJECTION_FAILURE_MESSAGE,
  });

/**
 * Projects the document snapshot and attributed review facts through one
 * bounded package scan. Optional review-part failures remain explicit unknown
 * fact families rather than invalidating the document snapshot.
 */
export const projectCompressedDocxWithReviewFacts = (
  bytes: Uint8Array,
  {
    textMaterialization = "word-host",
    highlightProjection = "resolved",
  }: ProjectCompressedDocxWithReviewFactsOptions = {},
): Promise<DocxPackageProjectionWire> =>
  projectWith({
    bytes,
    project: (input) => {
      if (textMaterialization === "word-host") {
        return projectCompressedDocxWithReviewFactsInWasm(input, highlightProjection);
      }
      return projectCompressedDocxWithReadableReviewFactsInWasm(input, highlightProjection);
    },
    message: PACKAGE_PROJECTION_FAILURE_MESSAGE,
  });
