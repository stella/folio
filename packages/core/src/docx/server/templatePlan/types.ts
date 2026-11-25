import { TaggedError, type TaggedErrorClass } from "better-result";

import type { FolioDocxConformanceIssue } from "../validateDocxConformance";

/** Why a template plan call stopped; every refusal names one of these. */
export const TEMPLATE_PLAN_ERROR_CODES: readonly [
  "archive-unreadable",
  "part-missing",
  "xml-unreadable",
  "plan-invalid",
  "sha-mismatch",
  "unknown-block",
  "hash-mismatch",
  "block-unplaced",
  "block-placed-twice",
  "strip-label-mismatch",
  "merge-invalid",
  "reference-span-invalid",
  "reference-target-unplaced",
  "highlight-span-invalid",
  "numbering-unknown",
  "style-unknown",
  "tracked-changes",
  "unsupported-content",
  "verification-failed",
] = Object.freeze([
  "archive-unreadable",
  "part-missing",
  "xml-unreadable",
  "plan-invalid",
  "sha-mismatch",
  "unknown-block",
  "hash-mismatch",
  "block-unplaced",
  "block-placed-twice",
  "strip-label-mismatch",
  "merge-invalid",
  "reference-span-invalid",
  "reference-target-unplaced",
  "highlight-span-invalid",
  "numbering-unknown",
  "style-unknown",
  "tracked-changes",
  "unsupported-content",
  "verification-failed",
] as const);

export type TemplatePlanErrorCode = (typeof TEMPLATE_PLAN_ERROR_CODES)[number];

const TemplatePlanErrorBase: TaggedErrorClass<"TemplatePlanError"> =
  TaggedError("TemplatePlanError");

/** A refused analysis, plan or output. `report` is set for `verification-failed`. */
export class TemplatePlanError extends TemplatePlanErrorBase<{
  message: string;
  code: TemplatePlanErrorCode;
  blockId?: string;
  part?: string;
  report?: TemplatePlanReport;
  cause?: unknown;
}> {}

export type DocxBlockNumbering = {
  numId: number;
  ilvl: number;
  abstractNumId: number;
};

export type DocxTableCell = {
  row: number;
  col: number;
  /** Ids of the blocks directly inside the cell, in order. */
  blockIds: string[];
};

/**
 * One block of the main story. `text` is the visible text: `\t` for a tab,
 * `\n` for a line break, U+FFFC for a footnote or endnote reference mark;
 * field results are included and field instructions are not. Offsets into
 * `text` are UTF-16 code units.
 */
export type DocxBlock = {
  /** `p<index>` for a paragraph, `t<index>` for a table. */
  id: string;
  /** Position among all blocks of the story in document order, nested ones included. */
  index: number;
  kind: "paragraph" | "table";
  text: string;
  /** The automatic list label as rendered, or null when the block is not numbered. */
  label: string | null;
  numbering: DocxBlockNumbering | null;
  styleId: string | null;
  styleName: string | null;
  outlineLevel: number | null;
  sectionIndex: number;
  /** Short hash of `text` and `label`; a plan pins a block with it. */
  hash: string;
  /** `w14:paraId` when the package states one. Informational. */
  paraId?: string;
  /** Drawings, pictures, embedded objects and equations anchored in the block. */
  inlineObjects: number;
  /** Table only: its cells in row order. */
  cells?: DocxTableCell[];
  /** Table only: every block inside the table in document order. */
  nested?: DocxBlock[];
};

export type DocxNote = { id: string; text: string };

export type DocxBlockManifest = {
  sha256: string;
  /** Top-level blocks of the main story in document order; a table is one block. */
  blocks: DocxBlock[];
  footnotes: DocxNote[];
  endnotes: DocxNote[];
  comments: number;
  /** Tracked insertions, deletions and moves in the main story and its notes. */
  trackedChanges: number;
  warnings: string[];
};

export type DocxTemplateStyle = {
  styleId: string;
  name: string | null;
  type: "paragraph" | "character" | "numbering" | "table";
  basedOn: string | null;
  /** The list the style binds its paragraphs to, inheritance resolved. */
  numPr: { numId: number; ilvl: number } | null;
  outlineLvl: number | null;
  isDefault: boolean;
};

export type DocxTemplateNumberingLevel = {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start: number;
  suff: "tab" | "space" | "nothing";
  pStyle: string | null;
};

export type DocxTemplateNumbering = {
  abstractNums: {
    abstractNumId: number;
    levels: DocxTemplateNumberingLevel[];
  }[];
  nums: {
    numId: number;
    abstractNumId: number;
    overrides: { ilvl: number; startOverride: number | null }[];
  }[];
};

export type DocxTemplateField = {
  /** Package path of the part holding the field. */
  part: string;
  instruction: string;
  /** The field name in upper case, e.g. `TOC`, `REF`, `FILLIN`, `DATE`. */
  kind: string;
  /** The cached result text. */
  result: string;
  /** Main story only: the top-level block the field starts in. */
  blockId?: string;
};

export type DocxTemplateSection = {
  index: number;
  orientation: "portrait" | "landscape";
  /** Page size and margins in twentieths of a point. */
  size: { width: number | null; height: number | null };
  margins: {
    top: number | null;
    right: number | null;
    bottom: number | null;
    left: number | null;
  };
  headers: { type: string; part: string | null }[];
  footers: { type: string; part: string | null }[];
};

export type DocxTemplateBookmark = {
  name: string;
  id: string;
  part: string;
};

export type DocxTemplateInspection = DocxBlockManifest & {
  styles: DocxTemplateStyle[];
  numbering: DocxTemplateNumbering;
  fields: DocxTemplateField[];
  sections: DocxTemplateSection[];
  bookmarks: DocxTemplateBookmark[];
};

export type TemplatePlanOmitReason =
  "empty" | "sourceToc" | "pageArtifact" | "duplicateLabel";

/**
 * The list a source paragraph joins.
 *
 * A number names a list of the template. `"new"` starts a list instance of
 * its own over the definition `baseNumId` names, or over the list the item's
 * style is bound to. Items sharing a `group` share the instance the first of
 * them starts; `restartFrom` is the number that instance begins with.
 */
export type TemplatePlanNumbering = {
  numId: number | "new";
  ilvl: number;
  restartFrom?: number;
  baseNumId?: number;
  group?: string;
};

export type TemplatePlanMerge = "lineBreak" | "space" | "none";

export type TemplatePlanNestedParagraph = {
  blockId: string;
  styleId?: string;
  numbering?: TemplatePlanNumbering | null;
  stripLabel?: string;
};

export type TemplatePlanSourceItem = {
  kind: "source";
  blockId: string;
  hash: string;
  /** A paragraph style of the template; for a table, a table style. */
  styleId?: string;
  /** `null` takes the paragraph out of the list its style would put it in. */
  numbering?: TemplatePlanNumbering | null;
  /** The manual label at the start of the block to remove, with the space after it. */
  stripLabel?: string;
  /** Join the next item of the plan, a source paragraph, into this paragraph. */
  mergeNext?: TemplatePlanMerge;
  /** Table only: the paragraph style of its cells. */
  cellStyleId?: string;
  /** Table only: placement of single paragraphs inside it. */
  nested?: TemplatePlanNestedParagraph[];
};

export type TemplatePlanItem =
  | { kind: "template"; templateBlockIds: string[] }
  | TemplatePlanSourceItem
  | { kind: "slot"; templateBlockId: string; sourceBlockIds: string[] };

export type TemplatePlanNumberingPatch = {
  abstractNumId: number;
  ilvl: number;
  numFmt?: string;
  lvlText?: string;
  start?: number;
  suff?: "tab" | "space" | "nothing";
};

export type TemplatePlanReference = {
  blockId: string;
  /** The span of the block's text the field replaces. */
  start: number;
  end: number;
  targetBlockId: string;
  /** The switches of the REF field, e.g. `\n \t \h`. */
  instruction: string;
  /** The cached result the field is written with. */
  expectedText: string;
};

export type TemplatePlanHighlight = {
  blockId: string;
  start: number;
  end: number;
  color: string;
  note?: string;
};

/** A plan carries block ids, enums and spans; the text comes from the source. */
export type TemplatePlan = {
  version: 1;
  sourceSha256: string;
  templateSha256: string;
  /** The new main story in order. */
  body: TemplatePlanItem[];
  omit: { blockId: string; reason: TemplatePlanOmitReason }[];
  /** Applied to the output's copy of the template numbering. */
  numberingPatches?: TemplatePlanNumberingPatch[];
  references?: TemplatePlanReference[];
  highlights?: TemplatePlanHighlight[];
  fields?: { freezeKinds?: string[]; updateOnOpen?: boolean };
};

export type TemplatePlanBlockResult = {
  /** The source block. */
  blockId: string;
  /** The output block holding its text; null when the output has none. */
  outputBlockId: string | null;
  disposition: "placed" | "merged" | "slot" | "nested" | "omitted";
  text: "identical" | "mismatch" | "omitted";
  emphasis: "identical" | "mismatch" | "omitted";
  /** What `stripLabel` removed, the following space included. */
  strippedPrefix?: string;
  /** The label the source showed: its automatic label, else the stripped one. */
  sourceLabel: string | null;
  /** The automatic label the output paragraph renders. */
  outputLabel: string | null;
  omitReason?: TemplatePlanOmitReason;
};

export type TemplatePlanReferenceResult = {
  blockId: string;
  targetBlockId: string;
  bookmark: string;
  instruction: string;
  expectedText: string;
  /** The result text the output field carries, or null when the field is missing. */
  cachedText: string | null;
  /** The source text the field replaced. */
  replacedText: string;
  /**
   * What the target's rendered label gives for a `\n` field, or the target's
   * text for a field without a number switch; null when it is not derived.
   */
  predictedText: string | null;
  status: "ok" | "missing-field" | "missing-bookmark" | "cache-mismatch";
};

export type TemplatePlanSpanCounts = {
  /** Spans in every block of the source. */
  source: number;
  /** Spans the output must show after the plan's strips, merges and references. */
  expected: number;
  /** Spans the output shows. */
  output: number;
};

export type TemplatePlanFieldLeftover = {
  part: string;
  kind: string;
  count: number;
};

export type TemplatePlanIssue = {
  code:
    | "text-mismatch"
    | "emphasis-mismatch"
    | "note-mismatch"
    | "block-missing"
    | "reference-broken"
    | "bookmark-dangling"
    | "field-unbalanced"
    | "relationship-dangling"
    | "note-dangling"
    | "numbering-dangling"
    | "style-dangling"
    | "prompting-field-left"
    | "update-fields-mismatch"
    | "conformance";
  message: string;
  blockId?: string;
  part?: string;
};

export type TemplatePlanReport = {
  version: 1;
  /** True when no issue was found; an output is only returned when it is. */
  ok: boolean;
  sourceSha256: string;
  templateSha256: string;
  outputSha256: string;
  /** One entry per block of the source, nested ones included, in document order. */
  blocks: TemplatePlanBlockResult[];
  counts: {
    sourceBlocks: number;
    placed: number;
    merged: number;
    omitted: number;
    templateBlocksKept: number;
    textIdentical: number;
    textMismatch: number;
    emphasisMismatch: number;
    /** Placed blocks whose source label and output label differ. */
    labelChanged: number;
    /** Placed blocks that showed a label in the source and render none. */
    labelLost: number;
  };
  footnotes: {
    source: number;
    output: number;
    identical: number;
    mismatch: number;
  };
  endnotes: {
    source: number;
    output: number;
    identical: number;
    mismatch: number;
  };
  emphasisSpans: TemplatePlanSpanCounts;
  highlightSpans: TemplatePlanSpanCounts;
  references: {
    planned: number;
    ok: number;
    /** Fields whose cached text differs from the text they replaced. */
    textChanging: number;
    /** Fields whose cached text differs from the derived prediction. */
    predictionMismatch: number;
    results: TemplatePlanReferenceResult[];
  };
  /** What the template brought along and the plan did not replace. */
  leftovers: {
    /** Fields still live in the output, by part and kind. */
    fields: TemplatePlanFieldLeftover[];
    /** Fields the plan froze to their cached text, by part and kind. */
    frozenFields: TemplatePlanFieldLeftover[];
    /** The filler strings the caller named that the output still contains. */
    filler: { text: string; part: string; count: number }[];
    /** Template blocks the plan kept verbatim. */
    templateBlockIds: string[];
  };
  dropped: {
    comments: number;
    sourceFields: { kind: string; count: number }[];
    internalHyperlinks: number;
    sectionBreaks: number;
  };
  updateFields: "set" | "cleared" | "unchanged";
  conformance: {
    status: "conformant" | "invalid" | "indeterminate";
    issues: FolioDocxConformanceIssue[];
  };
  issues: TemplatePlanIssue[];
  warnings: string[];
};

export type ApplyTemplatePlanOptions = {
  /** Template placeholder strings the report counts in the output. */
  fillerStrings?: readonly string[];
};

export type ApplyTemplatePlanArgs = {
  source: Uint8Array;
  template: Uint8Array;
  plan: TemplatePlan;
  options?: ApplyTemplatePlanOptions;
};

export type ApplyTemplatePlanResult = {
  docx: Uint8Array;
  report: TemplatePlanReport;
};

export type VerifyTemplatePlanOutputArgs = {
  source: Uint8Array;
  template: Uint8Array;
  output: Uint8Array;
  plan: TemplatePlan;
  options?: ApplyTemplatePlanOptions;
};
