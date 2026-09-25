/**
 * Shared primitives for ProseMirror to FlowBlock conversion: the conversion
 * options threaded through every converter, block-id allocation, and unit
 * conversion.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { FontAlternates } from "../../fonts/fontAlternates";
import type { ParagraphBlock, TableBlock, ParagraphAttrs } from "../../layout-engine/types";
import type { RunStyleResolver } from "../../prosemirror/runStyleFormatting";
import type { ListCounterStreams } from "../../prosemirror/listMarker";
import type { Theme, StyleDefinitions } from "../../types/document";

/**
 * Options for the conversion.
 */
export type ToFlowBlocksOptions = {
  /** Default font family. */
  defaultFont?: string;
  /** Default font size in points. */
  defaultSize?: number;
  /** Theme for resolving theme colors. */
  theme?: Theme | null;
  /** Document styles used to resolve character-style toggles without per-run copies. */
  styles?: StyleDefinitions;
  /** Document-scoped OOXML primary-font to `w:altName` lookup. */
  fontAlternates?: FontAlternates;
  /** Page content height in pixels (pageHeight - marginTop - marginBottom). Images taller than this are scaled down to fit. */
  pageContentHeight?: number;
  /** Whether a trailing page break moves its paragraph mark to the next page. */
  splitPageBreakAndParagraphMark?: boolean;
  /** Shared list counters for nested containers. */
  listCounters?: Map<number, number[]>;
  /** Latest concrete counters by abstract numbering definition. */
  listAbstractCounters?: Map<number, number[]>;
  /** Shared startOverride state for nested containers. */
  listSeenNumIds?: Set<string>;
  /**
   * Parallel counter state for the "original" (pre-revision) document, used to
   * number tracked-deletion list items. Word numbers inserted and deleted list
   * runs as if they never coexist: insertions get final-document numbering,
   * deletions keep their original numbering. Without a separate stream a deleted
   * item continues the counter of the inserted item before it (a, b → c, d, e
   * instead of a, b and a, b, c). Normal items advance both streams.
   */
  originalListCounters?: Map<number, number[]>;
  /** Latest concrete original-stream counters by abstract numbering definition. */
  originalListAbstractCounters?: Map<number, number[]>;
  /** Original-stream startOverride state. */
  originalListSeenNumIds?: Set<string>;
  /**
   * Document-wide `w:defaultTabStop` (§17.6.13) in twips. Stamped onto
   * every paragraph block so paragraph-local layout helpers (list marker
   * tab-stop math) can read it without taking a `Document` reference.
   * Defaults to the OOXML 720-twip value when absent.
   */
  defaultTabStopTwips?: number;
  /** Document-wide custom Word line-breaking settings. */
  lineBreakRules?: {
    noLineBreaksBefore?: { language?: string; characters: string };
    noLineBreaksAfter?: { language?: string; characters: string };
    useLegacyEthiopicAmharicRules?: boolean;
  };
  /** Document-generation policy for justified line fitting. */
  justificationCompatibility?: NonNullable<ParagraphAttrs["justificationCompatibility"]>;
  /** Document-generation policy for where `w:tblInd` is measured from. */
  tableIndentCompatibility?: NonNullable<TableBlock["indentCompatibility"]>;
  /**
   * When set, every table-cell anchor lays out inside its cell regardless of
   * an authored `wp:anchor/@layoutInCell="0"`. Compatibility mode 15 and
   * above ignore that opt-out; see `resolveAnchorLayoutInCellCompatibility`.
   */
  forceAnchorLayoutInCell?: boolean;
  /** Document-wide automatic hyphenation policy. */
  automaticHyphenation?: NonNullable<ParagraphAttrs["automaticHyphenation"]>;
  /** Line pitch for the final body section, whose properties live outside the PM body. */
  finalSectionDocumentGridLinePitchTwips?: number;
  /**
   * The number a note story's `w:footnoteRef`/`w:endnoteRef` marks show. The
   * mark stands for its note's reference number and shows it where it sits
   * among the story's runs, in that run's formatting.
   */
  noteReferenceMarkText?: string;
  /**
   * Endnotes laid out after the body's last block (`w:pos="docEnd"`). A body
   * that references one ends in its endnote area, so its final paragraph is no
   * longer the document's last and keeps its height.
   */
  trailingEndnoteIds?: ReadonlySet<number>;
};

export type FlowConversionOptions = ToFlowBlocksOptions & {
  firstPageBreakRunPosition: (node: PMNode) => number | undefined;
  /** `w:doNotUseIndentAsNumberingTabStop`, read from the document node. */
  numberingTabIgnoresIndent: boolean;
  listCounterStreams: ListCounterStreams;
  numberedRefResults?: ReadonlyMap<PMNode, string>;
  textBoxAnchorBlockIds: Map<string, ParagraphBlock["id"]>;
  styleResolver: RunStyleResolver;
};

/**
 * Convert twips to pixels (1 twip = 1/1440 inch, 1 inch = 96 CSS px).
 * No rounding — precision prevents cumulative layout drift across paragraphs.
 */
export function twipsToPixels(twips: number): number {
  return (twips / 1440) * 96;
}

/**
 * Generate a unique block ID.
 */
let blockIdCounter = 0;
export function nextBlockId(): string {
  return `block-${++blockIdCounter}`;
}

/**
 * Reset the block ID counter (useful for testing).
 */
export function resetBlockIdCounter(): void {
  blockIdCounter = 0;
}
