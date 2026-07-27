/* tslint:disable */
/* eslint-disable */

export type DocxProjectionFormattingSpan = readonly [
startUtf16: number,
endUtf16: number,
style: "bold" | "highlight",
];
export type DocxProjectionStructure =
| readonly []
| readonly [
type: "table",
tableId: string,
row: number,
column: number,
];
export type DocxProjectionParagraph = readonly [
ordinal: number,
text: string,
packageParagraphId: string | null,
formatting: readonly DocxProjectionFormattingSpan[],
structure: DocxProjectionStructure,
];
export type DocxProjectionFactSet<T> =
| readonly [status: "known", items: readonly T[]]
| readonly [status: "unknown", reason: DocxProjectionUnknownReason];
export type DocxProjectionUnknownReason =
| "document-part-only"
| "styles-part-unavailable"
| "unsupported-styles"
| "unsupported-numbering"
| "incomplete-bookmark-ranges"
| "unsupported-internal-references";
export type DocxProjectionIndentationFact = readonly [
paragraphOrdinal: number,
firstLineTwips: number | null,
hangingTwips: number | null,
leftTwips: number | null,
rightTwips: number | null,
startTwips: number | null,
endTwips: number | null,
firstLineCharsHundredths: number | null,
hangingCharsHundredths: number | null,
leftCharsHundredths: number | null,
rightCharsHundredths: number | null,
startCharsHundredths: number | null,
endCharsHundredths: number | null,
];
export type DocxProjectionNumberingFact = readonly [
paragraphOrdinal: number,
parentParagraphOrdinal: number | null,
childParagraphOrdinals: readonly number[],
];
export type DocxProjectionBookmarkFact = readonly [
paragraphOrdinal: number,
bookmarkId: number,
name: string,
span: DocxProjectionStructuralSpan,
];
export type DocxProjectionReferenceFact = readonly [
paragraphOrdinal: number,
referenceId: string,
role: "source" | "target",
span: DocxProjectionStructuralSpan,
];
export type DocxProjectionStructuralSpan = readonly [
startUtf8: number,
endUtf8: number,
startUtf16: number,
endUtf16: number,
coverage:
| "complete"
| "continues-before"
| "continues-after"
| "continues-before-and-after",
];
export type DocxProjectionStructuralFacts = readonly [
indentation: DocxProjectionFactSet<DocxProjectionIndentationFact>,
numberingHierarchy: DocxProjectionFactSet<DocxProjectionNumberingFact>,
bookmarks: DocxProjectionFactSet<DocxProjectionBookmarkFact>,
internalReferences: DocxProjectionFactSet<DocxProjectionReferenceFact>,
];
export type DocxProjectionRevisionUnsupportedReason =
| "incompatible-paragraph-merge"
| "structural-table-revision"
| "unsupported-revision-markup";
export type DocxProjectionRevisionStatus =
| readonly [status: "complete"]
| readonly [
status: "incomplete",
reasons: readonly DocxProjectionRevisionUnsupportedReason[],
];
export type DocxProjectionWire = readonly [
schemaVersion: 2,
paragraphs: readonly DocxProjectionParagraph[],
structuralFacts: DocxProjectionStructuralFacts,
revisionStatus: DocxProjectionRevisionStatus,
];



/**
 * Projects compressed DOCX bytes into a versioned host-independent snapshot.
 *
 * The ordinal is the paragraph's position in this immutable package snapshot.
 * Package `w14:paraId` values are returned separately. Neither value is an
 * application identity or a host navigation identity. A structural fact
 * family is `unknown` when the package does not prove completeness; an empty
 * `known` list is authoritative negative evidence.
 *
 * # Errors
 *
 * Returns a JavaScript `Error` when the package cannot be projected or a
 * numeric wire value cannot be represented by the schema.
 */
export function projectCompressedDocx(bytes: Uint8Array): DocxProjectionWire;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly projectCompressedDocx: (a: number, b: number, c: number) => void;
    readonly __wbindgen_add_to_stack_pointer: (a: number) => number;
    readonly __wbindgen_export: (a: number, b: number) => number;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
