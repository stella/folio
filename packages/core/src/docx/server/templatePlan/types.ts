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
  "label-not-writable",
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
  "label-not-writable",
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
  /**
   * The level prints a bullet: its label is a mark, whatever character
   * stands for it (a font's `o` as much as a dot), and no designator.
   */
  bullet?: true;
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
  /**
   * Page breaks among the paragraph's runs, when it holds any. A paragraph
   * without text can hold one: it is then all that starts the next page.
   */
  pageBreaks?: number;
  /**
   * The paragraph starts on a new page by its own formatting or its
   * style's (`w:pageBreakBefore`). A plan carries it over with an item's
   * `pageBreakBefore`; no other direct formatting of a paragraph comes along.
   */
  pageBreakBefore?: true;
  /**
   * The paragraph is the last of its section: the block after it opens
   * `sections[sectionIndex + 1]` of the manifest, whose `start` says whether
   * on a new page.
   */
  endsSection?: true;
  /**
   * The paragraph alignment in effect (`w:jc`, as written: `both`, `center`,
   * ...): the paragraph's own, else its style's with inheritance resolved.
   * Absent when neither states one.
   */
  alignment?: string;
  /** Table only: its cells in row order. */
  cells?: DocxTableCell[];
  /** Table only: every block inside the table in document order. */
  nested?: DocxBlock[];
};

export type DocxNote = { id: string; text: string };

/** Compatibility settings of a document that change how its text is laid out. */
export type DocxCompat = {
  /**
   * `w:doNotExpandShiftReturn`: a justified line that ends in a manual line
   * break keeps its spaces instead of being stretched to the margin.
   */
  doNotExpandShiftReturn: boolean;
};

/** One section of a document's main story: its page, and how it begins. */
export type DocxSection = {
  index: number;
  /**
   * How the section begins after the one before it: on the next page, an
   * even or odd one, in the next column, or on the same page. The first
   * section's value says nothing.
   */
  start: "nextPage" | "evenPage" | "oddPage" | "nextColumn" | "continuous";
  orientation: "portrait" | "landscape";
  /** Page size and margins in twentieths of a point; null where unstated. */
  size: { width: number | null; height: number | null };
  margins: {
    top: number | null;
    right: number | null;
    bottom: number | null;
    left: number | null;
  };
  /** Text columns of the section. */
  columns: number;
};

export type DocxBlockManifest = {
  sha256: string;
  /** Top-level blocks of the main story in document order; a table is one block. */
  blocks: DocxBlock[];
  /**
   * The sections of the main story in order; a block names its own by
   * `sectionIndex`. A body that states no section properties of its own
   * still ends in a section: it is listed with the defaults, nothing stated.
   */
  sections: DocxSection[];
  /** Compatibility settings of the document that change how its text is laid out. */
  compat: DocxCompat;
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
  /**
   * The paragraph alignment the style states (`w:jc`), inheritance resolved;
   * null when neither the style nor one it is based on states one.
   */
  alignment: string | null;
  /**
   * Whether the style sets its text in bold (`w:b`), inheritance and the
   * document's defaults resolved.
   */
  bold: boolean;
  /**
   * The font size the style gives its text, in half-points (`w:sz`),
   * inheritance and the document's defaults resolved; null when none of
   * them states one.
   */
  size: number | null;
  /**
   * Whether the style keeps its paragraph on the page of the next one
   * (`w:keepNext`), inheritance and the document's defaults resolved.
   */
  keepNext: boolean;
  isDefault: boolean;
};

export type DocxTemplateNumberingLevel = {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start: number;
  suff: "tab" | "space" | "nothing";
  pStyle: string | null;
  /** The level's paragraph indent in twentieths of a point; null where unstated. */
  indent: {
    left: number | null;
    hanging: number | null;
    firstLine: number | null;
  };
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
  /**
   * Main story only: the top-level block the field ends in. A plan keeps
   * every block from `blockId` to this one, or none of them.
   */
  endBlockId?: string;
};

/** A section of a template, with the header and footer parts it names. */
export type DocxTemplateSection = DocxSection & {
  headers: { type: string; part: string | null }[];
  footers: { type: string; part: string | null }[];
};

/**
 * A text effect of the `w14` extension. One whose values show nothing can still change
 * how a consumer treats the text: a run with a 3-D scene, even the default
 * one, is drawn as a picture.
 */
export type DocxTextEffect =
  | "glow"
  | "shadow"
  | "reflection"
  | "textOutline"
  | "textFill"
  | "scene3d"
  | "props3d";

/** What the run properties an effect stands in belong to. */
export type DocxTextEffectOwner =
  | { kind: "level"; abstractNumId: number | null; ilvl: number | null }
  | { kind: "style"; styleId: string | null }
  | { kind: "defaults" }
  | { kind: "text" };

/** A text effect of a template whose values show nothing. */
export type DocxTemplateInvisibleEffect = {
  part: string;
  effect: DocxTextEffect;
  owner: DocxTextEffectOwner;
};

export type DocxTemplateBookmark = {
  name: string;
  id: string;
  part: string;
};

export type DocxTemplateInspection = Omit<DocxBlockManifest, "sections"> & {
  styles: DocxTemplateStyle[];
  numbering: DocxTemplateNumbering;
  fields: DocxTemplateField[];
  sections: DocxTemplateSection[];
  bookmarks: DocxTemplateBookmark[];
  /**
   * Text effects that show nothing, in the document, its numbering, its
   * styles, its headers and its footers. A plan's output is written without
   * them unless the plan says otherwise.
   */
  invisibleEffects: DocxTemplateInvisibleEffect[];
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

/**
 * A paragraph's own indent, in twentieths of a point (`w:ind`). A side left
 * out stays what the style or the list level gives; a hanging indent and a
 * first-line indent exclude each other.
 */
export type TemplatePlanIndent = {
  left?: number;
  hanging?: number;
  firstLine?: number;
};

export type TemplatePlanNestedParagraph = {
  blockId: string;
  styleId?: string;
  numbering?: TemplatePlanNumbering | null;
  stripLabel?: string;
  /** As on a source item. */
  labelAsText?: boolean;
  /** As on a source item. */
  keepNext?: boolean;
  /** As on a source item. */
  indent?: TemplatePlanIndent;
};

/** A page other than the template's, in twentieths of a point. */
export type TemplatePlanSectionPage = {
  orientation: "portrait" | "landscape";
  width: number;
  height: number;
  /** Left out, the template's margins stay. */
  margins?: { top: number; right: number; bottom: number; left: number };
};

/**
 * A section of the output. It is built from the template's last section:
 * the same headers, footers and every other property, with only the page
 * given here put in its place. The first-page header and a page-number
 * start stay with the document's first section alone.
 */
export type TemplatePlanSection = {
  page?: TemplatePlanSectionPage;
  /**
   * How the section begins after the one before it; `nextPage` when left
   * out. The document's first section begins the document.
   */
  start?: "nextPage" | "evenPage" | "oddPage";
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
  /**
   * Write the block's automatic label as text at the start of its text, with
   * the tab or space its list puts after it: for a label no list of the
   * template prints. The paragraph it leads must join no list
   * (`numbering: null`). Only a label that reads the same wherever the file
   * is opened can be written: numbers, letters and roman numerals, not a
   * bullet glyph and not a number a language spells.
   */
  labelAsText?: boolean;
  /** Join the next item of the plan, a source paragraph, into this paragraph. */
  mergeNext?: TemplatePlanMerge;
  /**
   * Start the paragraph on a new page (`w:pageBreakBefore`): for a page the
   * source began by a break the plan does not place, by the paragraph's own
   * formatting, or by a section. Not for a table, and not for a block that
   * joins the paragraph before it.
   *
   * `false` is the same as leaving it out: nothing is written, and a style
   * of the template that starts its paragraphs on a new page still does.
   * Verification compares what the paragraph states for itself.
   */
  pageBreakBefore?: boolean;
  /**
   * Keep the paragraph on the page of the one after it (`w:keepNext`): for
   * a line that heads what follows in a style that does not say so. Not
   * for a table, and not for a block that joins the paragraph before it.
   * `false` is the same as leaving it out.
   */
  keepNext?: boolean;
  /**
   * The paragraph's own indent, over its style's and its list level's:
   * for a number written as text that hangs the way the list's would. Not
   * for a table, and not for a block that joins the paragraph before it.
   */
  indent?: TemplatePlanIndent;
  /**
   * The paragraph is the last of a section, described here: for a part of
   * the document whose page is another than the template's, such as a
   * landscape annex. Not for a table, a nested paragraph, or a block that
   * joins the paragraph before it. The template needs a last section of
   * its own to build from.
   */
  endsSection?: TemplatePlanSection;
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
  /**
   * The paragraph indent the level gives, in twentieths of a point. A value
   * left out keeps what the level states; `hanging` and `firstLine` exclude
   * each other, so stating one drops the other.
   */
  indent?: { left?: number; hanging?: number; firstLine?: number };
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
  /**
   * Compatibility settings to switch on in the output's settings, which are
   * otherwise the template's. `true` sets the setting at its place in the
   * schema's order; nothing else of the settings changes, and a setting is
   * never switched off.
   */
  compat?: { doNotExpandShiftReturn?: boolean };
  /**
   * The section the body ends with, when its page is another than the
   * template's. Left out, the template's last section closes the body as it
   * stands.
   */
  lastSection?: TemplatePlanSection;
  /**
   * What becomes of text effects whose values show nothing, in the parts
   * the output takes from the template: its document, numbering, styles,
   * headers and footers. `normalize`, the default, leaves them out, so text
   * that looked plain is written as text everywhere; an effect that shows
   * anything is never touched. `keep` writes those parts as they are.
   */
  textEffects?: "normalize" | "keep";
};

export type TemplatePlanEffectCount = {
  part: string;
  effect: DocxTextEffect;
  count: number;
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
  /** The automatic label `labelAsText` wrote as text, without what follows it. */
  labelAsText?: string;
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

export type TemplatePlanNoteCounts = {
  /** Notes in the source. */
  source: number;
  /** Notes the placed blocks refer to. */
  expected: number;
  /** Notes the output's placed blocks refer to. */
  output: number;
  identical: number;
  mismatch: number;
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
    | "style-mismatch"
    | "numbering-mismatch"
    | "page-start-mismatch"
    | "paragraph-format-mismatch"
    | "section-mismatch"
    | "template-mismatch"
    | "prompting-field-left"
    | "field-not-frozen"
    | "update-fields-mismatch"
    | "compat-mismatch"
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
    /** Placed blocks that showed a label in the source and show none: neither rendered nor written as text. */
    labelLost: number;
    /** Placed blocks whose automatic label is written as text. */
    labelAsText: number;
    /** Placed paragraphs the plan starts on a new page. */
    pageStarts: number;
    /** Placed paragraphs the plan ends a section with. */
    sectionEnds: number;
    /**
     * Empty paragraphs written between two tables the plan puts next to each
     * other, and after a table that ends the body: without one the tables
     * are one table, and a body cannot end in a table.
     */
    tableSeparators: number;
  };
  footnotes: TemplatePlanNoteCounts;
  endnotes: TemplatePlanNoteCounts;
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
  /**
   * Text effects that show nothing: those of the template the output does
   * not carry, and those it still does.
   */
  textEffects: {
    removed: TemplatePlanEffectCount[];
    left: TemplatePlanEffectCount[];
  };
  updateFields: "set" | "cleared" | "unchanged";
  conformance: {
    status: "conformant" | "invalid" | "indeterminate";
    /** The template's own status; the output may not be worse. */
    templateStatus: "conformant" | "invalid" | "indeterminate";
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
