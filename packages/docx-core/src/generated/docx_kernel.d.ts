/* tslint:disable */
/* eslint-disable */

export type DocxProjectionFormattingSpan = readonly [
startUtf16: number,
endUtf16: number,
style: "bold" | "highlight" | "superscript",
];
export type DocxProjectionStructure =
| readonly []
| readonly [
type: "table",
tableId: string,
row: number,
column: number,
];
export type DocxProjectionAlignmentValue = "center" | "justify" | "left" | "right";
export type DocxProjectionAlignmentSource = "direct" | "style";
export type DocxProjectionAlignment = readonly [
value: DocxProjectionAlignmentValue,
source: DocxProjectionAlignmentSource,
] | null;
export type DocxProjectionParagraph = readonly [
ordinal: number,
text: string,
packageParagraphId: string | null,
formatting: readonly DocxProjectionFormattingSpan[],
structure: DocxProjectionStructure,
styleId: string | null,
alignment: DocxProjectionAlignment,
];
export type DocxProjectionFormattingUnknownReason =
| "document-part-only"
| "styles-part-unavailable"
| "unsupported-styles";
export type DocxProjectionFormattingStatus =
| readonly [status: "complete"]
| readonly [
status: "incomplete",
reason: DocxProjectionFormattingUnknownReason,
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
export type DocxProjectionOutlineLevelFact = readonly [
paragraphOrdinal: number,
outlineLevel: number,
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
outlineLevels: DocxProjectionFactSet<DocxProjectionOutlineLevelFact>,
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
schemaVersion: 5,
paragraphs: readonly DocxProjectionParagraph[],
structuralFacts: DocxProjectionStructuralFacts,
revisionStatus: DocxProjectionRevisionStatus,
formattingStatus: DocxProjectionFormattingStatus,
];
export type DocxReviewUnknownReason =
| "invalid-document"
| "invalid-comments"
| "invalid-comments-extended"
| "resource-limit"
| "unsupported-location";
export type DocxReviewFactSet<T> =
| readonly [status: "known", items: readonly T[]]
| readonly [status: "unknown", reason: DocxReviewUnknownReason];
export type DocxRevisionKind =
| "insertion"
| "deletion"
| "moveFrom"
| "moveTo"
| "cellIns"
| "cellDel"
| "cellMerge"
| "pPrChange"
| "rPrChange"
| "sectPrChange"
| "tblPrChange"
| "trPrChange"
| "tcPrChange"
| "tblGridChange"
| "customXmlDelRangeStart"
| "customXmlDelRangeEnd"
| "customXmlInsRangeStart"
| "customXmlInsRangeEnd"
| "customXmlMoveFromRangeStart"
| "customXmlMoveFromRangeEnd"
| "customXmlMoveToRangeStart"
| "customXmlMoveToRangeEnd";
/**
 * Attributed revision wire tuple. Known content is flattened into the item so
 * one review fact creates one JavaScript array at the WASM boundary.
 */
export type DocxAttributedRevision =
| readonly [
type: DocxRevisionKind,
author: string,
date: string | null,
revisionId: string | null,
contentStatus: "known",
startParagraphOrdinal: number,
startUtf8: number,
startUtf16: number,
endParagraphOrdinal: number,
endUtf8: number,
endUtf16: number,
text: string,
contentKind: "text" | "formatting-only",
]
| readonly [
type: DocxRevisionKind,
author: string,
date: string | null,
revisionId: string | null,
contentStatus: "unknown",
reason: DocxReviewUnknownReason,
];
/**
 * Attributed comment wire tuple. Known content is flattened into the item so
 * one review fact creates one JavaScript array at the WASM boundary.
 */
export type DocxAttributedComment =
| readonly [
commentId: string,
author: string,
initials: string | null,
date: string | null,
parentCommentId: string | null,
threadState: "open" | "resolved",
contentStatus: "known",
startParagraphOrdinal: number,
startUtf8: number,
startUtf16: number,
endParagraphOrdinal: number,
endUtf8: number,
endUtf16: number,
commentText: string,
referencedText: string,
]
| readonly [
commentId: string,
author: string,
initials: string | null,
date: string | null,
parentCommentId: string | null,
threadState: "open" | "resolved",
contentStatus: "unknown",
reason: DocxReviewUnknownReason,
];
/** Review-fact wire schema. A new tuple layout requires a schema-version bump. */
export type DocxReviewFactsWire = readonly [
schemaVersion: 2,
revisions: DocxReviewFactSet<DocxAttributedRevision>,
comments: DocxReviewFactSet<DocxAttributedComment>,
];
/** Fused package wire schema. A new tuple layout requires a schema-version bump. */
export type DocxPackageProjectionWire = readonly [
schemaVersion: 2,
document: DocxProjectionWire,
reviewFacts: DocxReviewFactsWire,
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

/**
 * Projects the same fused snapshot with controls normalized for readable text.
 *
 * # Errors
 *
 * Returns a JavaScript `Error` under the same conditions as
 * [`project_compressed_docx_with_review_facts`].
 */
export function projectCompressedDocxWithReadableReviewFacts(bytes: Uint8Array): DocxPackageProjectionWire;

/**
 * Projects the document snapshot and attributed review facts from one bounded
 * package-directory scan.
 *
 * # Errors
 *
 * Returns a JavaScript `Error` when the document package itself cannot be
 * projected. Invalid optional review parts are represented as unknown facts.
 */
export function projectCompressedDocxWithReviewFacts(bytes: Uint8Array): DocxPackageProjectionWire;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly projectCompressedDocx: (a: number, b: number, c: number) => void;
    readonly projectCompressedDocxWithReadableReviewFacts: (a: number, b: number, c: number) => void;
    readonly projectCompressedDocxWithReviewFacts: (a: number, b: number, c: number) => void;
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
