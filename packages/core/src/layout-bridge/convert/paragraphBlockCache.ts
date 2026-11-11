/**
 * Paragraph blocks kept across the conversions of one editor's document.
 *
 * Converting a paragraph reads the node, the conversion constants, the marker
 * the list counters give it, and its position. An unchanged paragraph is the
 * same ProseMirror node in the next document, so a block converted from that
 * node at the same marker is the block converting it again would give, moved
 * by however far the paragraph moved. Its id is the one the conversion order
 * hands out where it sits now.
 */

import type { Node as PMNode } from "prosemirror-model";

import { documentFontFactsCacheIdentity } from "../../fonts/documentFontFacts";
import { noteParagraphOrigin } from "../../layout-engine/measure/paragraphKeyMemo";
import type { ParagraphBlock } from "../../layout-engine/types";
import { shiftParagraphPmAnchors } from "../../layout-engine/types";
import {
  getParagraphFrame,
  setParagraphFrame,
} from "../../utils/paragraphFrame";
import type {
  FlowConversionOptions,
  ListCounterOptionKey,
  ToFlowBlocksOptions,
} from "./flowConversionShared";
import type { ParagraphListMarker } from "./paragraphAttrs";
import type { RecordedConversion } from "./topLevelConversion";

type ConstantKey = Exclude<
  keyof ToFlowBlocksOptions,
  ListCounterOptionKey | "paragraphBlocks"
>;

/** Maps holding the same values under the same keys. */
const sameEntries = <K, V>(
  left: ReadonlyMap<K, V> | undefined,
  right: ReadonlyMap<K, V> | undefined,
): boolean =>
  left === right ||
  (left !== undefined &&
    right !== undefined &&
    left.size === right.size &&
    [...left].every(
      ([key, value]) => right.has(key) && right.get(key) === value,
    ));

const sameMembers = <T>(
  left: ReadonlySet<T> | undefined,
  right: ReadonlySet<T> | undefined,
): boolean =>
  left === right ||
  (left !== undefined &&
    right !== undefined &&
    left.size === right.size &&
    [...left].every((member) => right.has(member)));

/** Objects holding the same values under the same keys, one level deep. */
const sameFields = <T extends object>(
  left: T | undefined,
  right: T | undefined,
): boolean => {
  if (left === right) return true;
  if (left === undefined || right === undefined) return false;
  const keys = Object.keys(left);
  return (
    keys.length === Object.keys(right).length &&
    keys.every(
      (key) =>
        Object.hasOwn(right, key) &&
        Reflect.get(left, key) === Reflect.get(right, key),
    )
  );
};

/**
 * Every conversion option a paragraph's block can depend on, and how two
 * of its values are compared. Counter state is the one order-dependent input
 * and enters only as the list marker; the cache itself is not an input. A
 * layout run builds some of these afresh from the same document, so they
 * compare by content.
 */
const CONVERSION_CONSTANTS: {
  readonly [Key in ConstantKey]-?: (
    left: ToFlowBlocksOptions,
    right: ToFlowBlocksOptions,
  ) => boolean;
} = {
  altChunkMarkerText: (left, right) =>
    left.altChunkMarkerText === right.altChunkMarkerText,
  altChunkContent: (left, right) =>
    left.altChunkContent === right.altChunkContent,
  showCommentHighlights: (left, right) =>
    left.showCommentHighlights === right.showCommentHighlights,
  embeddedFontFamilyMap: (left, right) =>
    left.embeddedFontFamilyMap === right.embeddedFontFamilyMap,
  documentFontFacts: (left, right) =>
    documentFontFactsCacheIdentity(left.documentFontFacts) ===
    documentFontFactsCacheIdentity(right.documentFontFacts),
  defaultFont: (left, right) => left.defaultFont === right.defaultFont,
  defaultSize: (left, right) => left.defaultSize === right.defaultSize,
  theme: (left, right) => left.theme === right.theme,
  styles: (left, right) => left.styles === right.styles,
  fontAlternates: (left, right) =>
    sameEntries(left.fontAlternates, right.fontAlternates),
  pageContentHeight: (left, right) =>
    left.pageContentHeight === right.pageContentHeight,
  splitPageBreakAndParagraphMark: (left, right) =>
    left.splitPageBreakAndParagraphMark ===
    right.splitPageBreakAndParagraphMark,
  defaultTabStopTwips: (left, right) =>
    left.defaultTabStopTwips === right.defaultTabStopTwips,
  lineBreakRules: (left, right) =>
    sameFields(left.lineBreakRules, right.lineBreakRules),
  justificationCompatibility: (left, right) =>
    sameFields(
      left.justificationCompatibility,
      right.justificationCompatibility,
    ),
  tableIndentCompatibility: (left, right) =>
    sameFields(left.tableIndentCompatibility, right.tableIndentCompatibility),
  forceAnchorLayoutInCell: (left, right) =>
    left.forceAnchorLayoutInCell === right.forceAnchorLayoutInCell,
  automaticHyphenation: (left, right) =>
    sameFields(left.automaticHyphenation, right.automaticHyphenation),
  finalSectionDocumentGridLinePitchTwips: (left, right) =>
    left.finalSectionDocumentGridLinePitchTwips ===
    right.finalSectionDocumentGridLinePitchTwips,
  noteReferenceMarkText: (left, right) =>
    left.noteReferenceMarkText === right.noteReferenceMarkText,
  trailingEndnoteIds: (left, right) =>
    sameMembers(left.trailingEndnoteIds, right.trailingEndnoteIds),
};

const CONSTANT_KEYS = Object.keys(
  CONVERSION_CONSTANTS,
) as readonly ConstantKey[];

type KeptParagraph = {
  listMarker: ParagraphListMarker;
  /** Where the paragraph started when it was converted. */
  pmStart: number;
  /** The block as converting it gave it, never handed out itself. */
  block: ParagraphBlock;
};

/**
 * A copy of `block` moved by `delta`, under `id`. The copy owns its top-level
 * fields, attributes and runs, which conversion stamps after a paragraph is
 * converted; the nested values it shares are never written after conversion.
 */
const movedBlock = (
  block: ParagraphBlock,
  delta: number,
  id: ParagraphBlock["id"],
): ParagraphBlock => {
  const moved: ParagraphBlock = {
    ...shiftParagraphPmAnchors(block, delta),
    id,
    attrs: { ...block.attrs },
  };
  // A text frame is kept beside the block, not on it: the copy carries it too.
  const frame = getParagraphFrame(block);
  if (frame !== undefined) setParagraphFrame(moved, frame);
  return moved;
};

/**
 * Whether converting `node` reads anything besides itself, the constants and
 * its list marker: a field may show another paragraph's number, and a text box
 * anchor records which block holds it.
 */
const readsOtherBlocks = (node: PMNode): boolean => {
  let reads = false;
  node.descendants((child) => {
    if (reads) return false;
    if (
      child.type.name === "field" ||
      child.type.name === "structuredField" ||
      child.type.name === "textBoxAnchor"
    ) {
      reads = true;
      return false;
    }
    return true;
  });
  return reads;
};

export class ParagraphBlockCache {
  private constants: FlowConversionOptions | undefined;
  private paragraphs = new WeakMap<PMNode, KeptParagraph | null>();
  private conversion: RecordedConversion | undefined;

  /** Start over unless every conversion constant converts as the one the kept blocks were converted with. */
  use(options: FlowConversionOptions): void {
    const kept = this.constants;
    this.constants = options;
    if (
      kept !== undefined &&
      kept.numberingTabIgnoresIndent === options.numberingTabIgnoresIndent &&
      CONSTANT_KEYS.every((key) => CONVERSION_CONSTANTS[key](kept, options))
    )
      return;
    this.paragraphs = new WeakMap();
    this.conversion = undefined;
  }

  /**
   * The last conversion, when it converted a document with the same document
   * attributes under the constants `use` was last given.
   */
  lastConversion(doc: PMNode): RecordedConversion | undefined {
    const conversion = this.conversion;
    return conversion !== undefined && conversion.doc.sameMarkup(doc)
      ? conversion
      : undefined;
  }

  /** Keep what converting each top-level node gave, for the next conversion. */
  recordConversion(conversion: RecordedConversion | undefined): void {
    this.conversion = conversion;
  }

  /** The block converting `node` at `pmStart` with `listMarker` gives, under `id`. */
  take(
    node: PMNode,
    pmStart: number,
    listMarker: ParagraphListMarker,
    id: () => ParagraphBlock["id"],
  ): ParagraphBlock | undefined {
    const kept = this.paragraphs.get(node);
    if (!kept || kept.listMarker !== listMarker) return undefined;
    const moved = movedBlock(kept.block, pmStart - kept.pmStart, id());
    noteParagraphOrigin(moved, kept.block);
    return moved;
  }

  /**
   * Keep the block converting `node` at `pmStart` gave, unless converting it
   * read other blocks. Call before anything else writes to `block`.
   */
  keep(
    node: PMNode,
    pmStart: number,
    listMarker: ParagraphListMarker,
    block: ParagraphBlock,
  ): void {
    if (this.paragraphs.get(node) === null) return;
    if (readsOtherBlocks(node)) {
      this.paragraphs.set(node, null);
      return;
    }
    const kept = movedBlock(block, 0, block.id);
    noteParagraphOrigin(block, kept);
    this.paragraphs.set(node, { listMarker, pmStart, block: kept });
  }
}
