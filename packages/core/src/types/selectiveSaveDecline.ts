import type { OpStory } from "@stll/docx-core/ops";

/** Why the structural splice declined, so the save repacked the package. */
export type StructuralSpliceDeclineReason =
  /** A body the splice cannot read, or one in Strict conformance. */
  | "structuralUnreadableBody"
  /** A paragraph without a valid unique paragraph id. */
  | "structuralParagraphIds"
  /** A body block other than a paragraph, table, section or range marker. */
  | "structuralBodyBlock"
  /** A changed paragraph holding a section, note reference or drawing. */
  | "structuralDependentContent"
  /** A table, body marker or the final section changed shape. */
  | "structuralBarrierChanged"
  /** Surviving paragraphs reordered, or a new paragraph reusing an id. */
  | "structuralReorder"
  /** A serialized paragraph whose namespaces the source cannot carry. */
  | "structuralFragment"
  /**
   * A range or field the splice would split: the spliced body would place a
   * bookmark, comment, permission, move or field marker differently from the
   * serialized body.
   */
  | "structuralRangeSplit"
  /** The splices overlap, or would leave a comment range unbalanced. */
  | "structuralSpliceRefused";
export type SelectiveSaveDeclineReason =
  | StructuralSpliceDeclineReason
  | "untrackedChanges"
  | "tooLarge"
  | "unmaterializedBodyResource"
  | "unmaterializedStoryResource"
  | "unmaterializedHeaderFooter"
  | "pictureWatermark"
  | "orphanCommentRanges"
  | "unsynthesizedReplyRanges"
  | "invalidModel"
  | "unmaterializedHyperlink"
  | "unsafePackagePath"
  | "missingDocumentPart"
  | "addedNumbering"
  | "unroutedParagraphs"
  | "paragraphPatch"
  | "commentsExtended"
  | "numberingPart"
  | "styleMaterialization"
  | "strictConformance"
  | "packageError";
/** The attempted selective write's refusal, without attributing package-wide failures to a part. */
export type SelectiveSaveDecline = { reason: SelectiveSaveDeclineReason } & (
  | { scope: "package"; part?: never; story?: never }
  | { scope: "part"; part: string; story?: never }
  | {
      scope: "story";
      story: Extract<OpStory, { kind: "header" | "footer" }>;
      part?: never;
    }
);
