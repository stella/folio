/**
 * Shared primitives for ProseMirror to FlowBlock conversion: the conversion
 * options threaded through every converter, block-id allocation, and unit
 * conversion.
 */

import type { Node as PMNode } from "prosemirror-model";

import type { DocumentFontFacts } from "../../fonts/documentFontFacts";
import type { FontAlternates } from "../../fonts/fontAlternates";
import type {
  ParagraphAttrs,
  ParagraphBlock,
  TableBlock,
} from "../../layout-engine/types";
import type { ListCounterStreams } from "../../prosemirror/listMarker";
import type { RunStyleResolver } from "../../prosemirror/runStyleFormatting";
import type { StyleDefinitions, Theme } from "../../types/document";
import type { ParagraphBlockCache } from "./paragraphBlockCache";

/**
 * Options for the conversion.
 */
/** How much of a `w:altChunk` a story has room to show. */
export const ALT_CHUNK_CONTENT: {
  readonly text: "text";
  readonly markerOnly: "markerOnly";
} = {
  /** Its own text when it has any, otherwise the marker. */
  text: "text",
  /** Only the marker: the story's height cannot follow the chunk's length. */
  markerOnly: "markerOnly",
} as const;
export type AltChunkContent =
  (typeof ALT_CHUNK_CONTENT)[keyof typeof ALT_CHUNK_CONTENT];

export type ToFlowBlocksOptions = {
  /** Localised text for a `w:altChunk` that has no text of its own to show. */
  altChunkMarkerText?: string;
  /** Defaults to the chunk's own text. */
  altChunkContent?: AltChunkContent;
  /** Paint comment anchors and highlights; false leaves authored marks intact. */
  showCommentHighlights?: boolean;
  /** Embedded families owned by this editor, never ambient process state. */
  embeddedFontFamilyMap?: ReadonlyMap<string, string>;
  /** Authored font-table facts owned by this document. */
  documentFontFacts?: DocumentFontFacts;
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
  justificationCompatibility?: NonNullable<
    ParagraphAttrs["justificationCompatibility"]
  >;
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
  /**
   * Paragraph blocks kept from earlier conversions of the same editor's
   * document. Conversion gives the same blocks with or without it.
   */
  paragraphBlocks?: ParagraphBlockCache;
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

/** The options holding list counter state, which each paragraph advances in story order. */
export type ListCounterOptionKey =
  | "listCounters"
  | "listAbstractCounters"
  | "listSeenNumIds"
  | "originalListCounters"
  | "originalListAbstractCounters"
  | "originalListSeenNumIds"
  | "listCounterStreams";

/** The conversion options without list counter state: what converting one paragraph at a known list marker reads. */
export type CounterFreeFlowConversionOptions = Omit<
  FlowConversionOptions,
  ListCounterOptionKey
>;

/**
 * Convert twips to pixels (1 twip = 1/1440 inch, 1 inch = 96 CSS px).
 * No rounding — precision prevents cumulative layout drift across paragraphs.
 */
export function twipsToPixels(twips: number): number {
  return (twips / 1440) * 96;
}

/**
 * Block ids of the conversion under way. Fresh ids count up from where the
 * conversion started; a node converting again first takes back, in order, the
 * ids it held in the conversion before, so a block keeps its id across passes.
 */
let blockIdCounter = 0;
let preferredBlockIds: readonly string[] = [];
let preferredAt = 0;
let handedOut: string[] | undefined;

/** A block id unique within the conversion under way. */
export function nextBlockId(): string {
  const id =
    preferredAt < preferredBlockIds.length
      ? (preferredBlockIds[preferredAt++] ?? `block-${++blockIdCounter}`)
      : `block-${++blockIdCounter}`;
  handedOut?.push(id);
  return id;
}

/**
 * Start a conversion whose fresh ids count up from `freshFrom`: above every
 * id an earlier conversion it reuses ids from handed out.
 */
export function resetBlockIdCounter(freshFrom = 0): void {
  blockIdCounter = freshFrom;
  preferredBlockIds = [];
  preferredAt = 0;
  handedOut = undefined;
}

/** Where fresh ids of the conversion under way have counted up to. */
export function freshBlockIdsFrom(): number {
  return blockIdCounter;
}

/**
 * Hand out `ids` first, in order, then fresh ids, and record every id handed
 * out until {@link takeHandedOutBlockIds}. For one node's conversion.
 */
export function preferBlockIds(ids: readonly string[]): void {
  preferredBlockIds = ids;
  preferredAt = 0;
  handedOut = [];
}

/** The ids handed out since {@link preferBlockIds}; stops recording. */
export function takeHandedOutBlockIds(): readonly string[] {
  const ids = handedOut ?? [];
  preferredBlockIds = [];
  preferredAt = 0;
  handedOut = undefined;
  return ids;
}
