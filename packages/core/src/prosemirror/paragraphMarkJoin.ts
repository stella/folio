/**
 * Removing a paragraph mark: the one join both a resolved revision and a
 * retracted one perform.
 *
 * A paragraph's mark ends it, and in OOXML its properties live on that mark.
 * With the mark gone, the paragraph runs on into the next one and ends with
 * the NEXT paragraph's mark, so the paragraph left is the next one: its
 * paragraph properties (style, numbering, the mark's own run properties), its
 * mark's revision and its section endpoint all survive, and the removed mark
 * takes its paragraph's properties and section endpoint with it. This holds
 * whether the first paragraph still has words or not.
 *
 * Its identity (`paraId`, `textId`) is the next one's too, unless the join
 * empties the next paragraph: accepting its deletion or rejecting its
 * insertion takes every word it held, so the reviewer removed it and its mark
 * stays only to end the container. The paragraph left then stands for the
 * first one, as it does applied directly (`joinIdentityOf`). The resolvers
 * and the tracked edits that join decide that one way (`joinedNextContentOf`).
 *
 * The first paragraph's runs keep their direct formatting; what its style
 * lent them is re-read in the next paragraph's style.
 */

import { panic } from "better-result";
import {
  Fragment,
  type MarkType,
  type Node as PMNode,
} from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import type { Mapping } from "prosemirror-transform";

import {
  getProseParagraphPropertySourceToken,
  joinProseParagraphsWithRightPropertySource,
  markParagraphPropertySourceTransfers,
  PARAGRAPH_SOURCE_JOIN_RULES,
} from "../docx/paragraphPropertySource";
import { paragraphJoinAfter } from "./commands/paragraphBookmarkJoin";
import { JOINED_RUNS_RESTYLED_META } from "./extensions/features/JoinedRunStyleExtension";
import { rebaseParagraphRuns } from "./rebaseParagraphRunFormatting";
import {
  paragraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";
import type { ParagraphAttrs, SuggestedStructuralMarker } from "./schema/nodes";
import { paragraphContentOf } from "./zeroWidthAnchors";

/**
 * What a join leaves of the next paragraph: its own content (`kept`), or
 * none, because the review removes every word it holds (`emptied`).
 */
export type JoinedNextContent = "kept" | "emptied";

/**
 * The revision marks of content still pending when a join runs: what this
 * direction of review removes (`deletion` accepting, `insertion` rejecting)
 * and what it keeps; null when nothing is pending any more.
 */
export type PendingContent = { removedBy: MarkType; keptBy: MarkType } | null;

type JoinedNextContentOptions = {
  /** The paragraph whose mark goes, as it stands. */
  first: PMNode;
  /**
   * The paragraphs the next one stood for before the resolution joining
   * them, in order; empty for a join that resolves nothing.
   */
  before: readonly PMNode[];
  /** The next paragraphs as they stand, in order. */
  after: readonly PMNode[];
  pending: PendingContent;
};

/**
 * Whether `paragraph` holds content other than what `without` marks, by the
 * engine's paragraph-content rule.
 */
const holdsContentWithout = (
  paragraph: PMNode,
  without: readonly MarkType[],
): boolean => {
  if (without.length === 0) return Boolean(paragraphContentOf(paragraph));
  const rest: PMNode[] = [];
  paragraph.forEach((child) => {
    if (!without.some((mark) => mark.isInSet(child.marks))) rest.push(child);
  });
  return Boolean(paragraphContentOf(paragraph.copy(Fragment.fromArray(rest))));
};

/**
 * Whether a join empties the next paragraph: it holds nothing that resolving
 * every pending revision this way keeps, and either held content before the
 * resolution joining them (the reviewer removed its words, so its mark stays
 * only to end the container) or the first paragraph holds words no revision
 * added or removes, which are the paragraph left whichever way the rest is
 * resolved. A paragraph that was empty all along, after one with nothing
 * settled, is kept.
 */
export const joinedNextContentOf = ({
  first,
  before,
  after,
  pending,
}: JoinedNextContentOptions): JoinedNextContent => {
  const removed = pending === null ? [] : [pending.removedBy];
  if (after.some((node) => holdsContentWithout(node, removed))) return "kept";
  const settled = pending === null ? [] : [pending.removedBy, pending.keptBy];
  return before.some((node) => paragraphContentOf(node)) ||
    holdsContentWithout(first, settled)
    ? "emptied"
    : "kept";
};

/**
 * What a join that tracked edits make, retracting an added break or retiring
 * the break before an emptied last paragraph, leaves of the next paragraph:
 * judged as accepting the edits would, with nothing resolved yet.
 */
export const trackedJoinNextContent = (
  paragraph: PMNode,
  next: PMNode,
): JoinedNextContent => {
  const { deletion, insertion } = paragraph.type.schema.marks;
  if (!deletion || !insertion)
    return panic("Tracked edits need insertion and deletion marks");
  return joinedNextContentOf({
    first: paragraph,
    before: [],
    after: [next],
    pending: { removedBy: deletion, keptBy: insertion },
  });
};

type SourceParagraphsOptions = {
  /** The document before resolving. */
  source: PMNode;
  /** Positions in `source` to positions in the document being resolved. */
  mapping: Mapping;
  /** A range of the document being resolved. */
  range: { from: number; to: number };
};

/**
 * The paragraphs of `source` that `range` stood for before resolving: what
 * `joinedNextContentOf` reads as `before`.
 */
export const sourceParagraphsOf = ({
  source,
  mapping,
  range,
}: SourceParagraphsOptions): PMNode[] => {
  const inverse = mapping.invert();
  const from = inverse.map(range.from, 1);
  const to = inverse.map(range.to, -1);
  const paragraphs: PMNode[] = [];
  if (from >= to) return paragraphs;
  source.nodesBetween(from, to, (node) => {
    if (!node.isTextblock) return true;
    paragraphs.push(node);
    return false;
  });
  return paragraphs;
};

/** The two paragraphs a join runs together, and what stands between them. */
type JoinSides = {
  /** The paragraph whose mark goes. */
  paragraph: PMNode;
  /** The paragraph after it, whose mark stays. */
  next: PMNode;
  /** Inline size placed between the two paragraphs' content (bookmarks). */
  between: number;
  /** What the join leaves of the next paragraph. */
  nextContent: JoinedNextContent;
};

/**
 * Which paragraph the joined one stands for: the next, whose mark stays,
 * unless the join empties it; then the first, whose words are all the joined
 * paragraph keeps.
 */
const joinIdentityOf = ({
  paragraph,
  next,
  nextContent,
}: JoinSides): PMNode => {
  switch (nextContent) {
    case "kept":
      return next;
    case "emptied":
      return paragraph;
  }
};

/**
 * Where one attribute of the joined paragraph comes from: `"next"` for the
 * mark that stays (its properties, its revision), or a rule that reads both
 * sides.
 */
type JoinedAttrRule = "next" | ((sides: JoinSides) => unknown);

const nextValue = (key: string) => (sides: JoinSides) => sides.next.attrs[key];
const identityValue = (key: string) => (sides: JoinSides) =>
  joinIdentityOf(sides).attrs[key];

/** The paragraph-level records each side holds, in document order. */
const bothSides =
  (key: string, shift: (item: unknown, by: number) => unknown) =>
  ({ paragraph, next, between }: JoinSides) => {
    const left: unknown = paragraph.attrs[key];
    const right: unknown = next.attrs[key];
    const items = [
      ...(Array.isArray(left) ? left : []),
      ...(Array.isArray(right)
        ? right.map((item: unknown) =>
            shift(item, paragraph.content.size + between),
          )
        : []),
    ];
    return items.length === 0 ? right : items;
  };

/** The suggestion a paragraph's whole-paragraph insert marker names, if any. */
const suggestedInsertOf = (
  node: PMNode,
): (SuggestedStructuralMarker & { suggestionId: string }) | null => {
  const marker: unknown = node.attrs["_suggestedInsert"];
  return typeof marker === "object" &&
    marker !== null &&
    typeof (marker as { suggestionId?: unknown }).suggestionId === "string" &&
    typeof (marker as { revisionId?: unknown }).revisionId === "number"
    ? (marker as SuggestedStructuralMarker & { suggestionId: string })
    : null;
};

/**
 * Whether the joined paragraph is still all one suggestion's insertion: only
 * when both paragraphs were. A paragraph that held other words is not, so it
 * keeps no whole-paragraph marker; the mark it ends with records the next
 * paragraph's suggested insertion instead (`pPrMark`).
 */
const stillWhollySuggested = ({ paragraph, next }: JoinSides): boolean => {
  const left = suggestedInsertOf(paragraph);
  const right = suggestedInsertOf(next);
  return (
    left !== null && right !== null && left.suggestionId === right.suggestionId
  );
};

/**
 * Every attribute of the joined paragraph, decided one by one. The mark that
 * stays is the next paragraph's, so most come from it; what the paragraph
 * holds as content (boundary bookmarks, empty hyperlinks) comes from both,
 * and a suggestion's whole-paragraph marker survives only where the whole
 * joined paragraph is still that suggestion's.
 */
const JOINED_PARAGRAPH_ATTRS = {
  // Identity: the paragraph the joined one stands for.
  paraId: identityValue("paraId"),
  textId: identityValue("textId"),
  idStability: identityValue("idStability"),
  ...PARAGRAPH_SOURCE_JOIN_RULES,
  reviewCarrier: "next",
  // The mark's paragraph properties.
  alignment: "next",
  alignmentFromStyle: "next",
  kinsoku: "next",
  autoSpaceDE: "next",
  autoSpaceDN: "next",
  overflowPunctuation: "next",
  suppressAutoHyphens: "next",
  spaceBefore: "next",
  spaceAfter: "next",
  lineSpacing: "next",
  lineSpacingRule: "next",
  lineSpacingExplicit: "next",
  snapToGrid: "next",
  spacingExplicit: "next",
  spacingFromDocDefaults: "next",
  spacingFromImplicitDefaultStyle: "next",
  indentLeft: "next",
  indentRight: "next",
  indentFirstLine: "next",
  hangingIndent: "next",
  numPr: "next",
  _pastedHtmlList: "next",
  numPrFromStyle: "next",
  listNumFmt: "next",
  listIsBullet: "next",
  listIsLegal: "next",
  listMarker: "next",
  listMarkerTemplate: "next",
  listMarkerHidden: "next",
  listMarkerFormatting: "next",
  listMarkerAlignment: "next",
  listMarkerSuffix: "next",
  listMarkerAllCaps: "next",
  listImplicitChildLevelAdvances: "next",
  listFoldedMarkerSuffix: "next",
  listMarkerSecondSlotOffsetTwips: "next",
  listLevelNumFmts: "next",
  listLevelStarts: "next",
  listLevelTabs: "next",
  listAbstractNumId: "next",
  listStartOverride: "next",
  styleId: "next",
  _tableOfContentsLevel: "next",
  borders: "next",
  shading: "next",
  tabs: "next",
  pageBreakBefore: "next",
  renderedPageBreakBefore: "next",
  keepNext: "next",
  keepLines: "next",
  widowControl: "next",
  contextualSpacing: "next",
  runInWithNext: "next",
  defaultTextFormatting: "next",
  _tableRunFormatting: "next",
  direction: "next",
  outlineLevel: "next",
  _originalFormatting: "next",
  _resolvedFormatting: "next",
  _styleResolvedFormatting: "next",
  _autospacingBase: "next",
  _detachedWatermarkHost: "next",
  // The section and the reviews the mark carries; callers move a section
  // the removed mark ended.
  _sectionProperties: "next",
  _propertyChanges: "next",
  // The words: a page break after the text is the next paragraph's last; a
  // paragraph holding only a page break is one only if this one held nothing.
  _trailingPageBreak: "next",
  _pageBreakCarrier: (sides) =>
    sides.paragraph.content.size === 0
      ? sides.next.attrs["_pageBreakCarrier"]
      : null,
  bookmarks: bothSides("bookmarks", (item) => item),
  _emptyHyperlinks: bothSides("_emptyHyperlinks", (item, by) =>
    typeof item === "object" && item !== null && "offset" in item
      ? { ...item, offset: Number(item.offset) + by }
      : item,
  ),
  // A split copies the paragraph's attrs onto both halves, and a save writes
  // the sessions (`w:rsid*`) on the first half only. Rejecting the split
  // removes that first half, so the paragraph left names none of them.
  _preservedAttributes: ({ paragraph, next }) => {
    const sessions: unknown = next.attrs["_preservedAttributes"];
    return typeof sessions === "object" &&
      sessions !== null &&
      sessions === paragraph.attrs["_preservedAttributes"]
      ? null
      : sessions;
  },
  // The next paragraph's inserted mark: a suggested whole paragraph's
  // marker becomes its suggested inserted mark once the joined paragraph
  // holds words that were not part of it.
  pPrMark: (sides) => {
    const own: unknown = sides.next.attrs["pPrMark"];
    if (own != null || stillWhollySuggested(sides)) return own;
    const suggested = suggestedInsertOf(sides.next);
    return suggested === null
      ? own
      : {
          kind: "ins",
          info: {
            id: suggested.revisionId,
            author: suggested.author,
            ...(suggested.date == null ? {} : { date: suggested.date }),
            ...(typeof suggested.initials === "string"
              ? { initials: suggested.initials }
              : {}),
            provenance: "suggested",
            suggestionId: suggested.suggestionId,
          },
        };
  },
  _suggestedInsert: (sides) =>
    stillWhollySuggested(sides) ? sides.next.attrs["_suggestedInsert"] : null,
} as const satisfies Record<
  | keyof ParagraphAttrs
  | keyof typeof PARAGRAPH_SOURCE_JOIN_RULES
  | "idStability"
  | "_detachedWatermarkHost",
  JoinedAttrRule
>;

/** The attributes a join decides, for a check that the schema adds none undecided. */
export const JOINED_PARAGRAPH_ATTR_NAMES: readonly string[] = Object.keys(
  JOINED_PARAGRAPH_ATTRS,
);

type JoinedParagraphAttrsOptions = {
  paragraph: PMNode;
  next: PMNode;
  /** The inline size the join places between the two paragraphs' content. */
  between?: number;
  nextContent: JoinedNextContent;
};

/**
 * The attrs of the paragraph a join leaves, each by its rule in
 * `JOINED_PARAGRAPH_ATTRS`.
 */
export const joinedParagraphAttrs = ({
  paragraph,
  next,
  between = 0,
  nextContent,
}: JoinedParagraphAttrsOptions): PMNode["attrs"] => {
  const sides = { paragraph, next, between, nextContent };
  const rules: Readonly<Record<string, JoinedAttrRule>> =
    JOINED_PARAGRAPH_ATTRS;
  const attrs: Record<string, unknown> = {};
  let changed = false;
  for (const key of Object.keys(next.attrs)) {
    const rule = rules[key];
    if (rule === undefined)
      panic(`A paragraph attribute has no join rule: ${key}`);
    const value = (rule === "next" ? nextValue(key) : rule)(sides);
    attrs[key] = value;
    if (value !== next.attrs[key]) changed = true;
  }
  return changed ? attrs : next.attrs;
};

type JoinAtParagraphMarkOptions = {
  tr: Transaction;
  /** Position of the paragraph whose mark goes. */
  paragraphPos: number;
  paragraph: PMNode;
  /** The paragraph directly after it, which the join runs into. */
  next: PMNode;
  /** What the join leaves of `next`: see `joinedNextContentOf`. */
  nextContent: JoinedNextContent;
  /**
   * Re-reads the moved runs' inherited formatting in the next paragraph's
   * style. Without one, the editor's join restyler does it when the
   * transaction is dispatched.
   */
  styleResolver: RunStyleResolver | null;
};

export const joinAtParagraphMark = ({
  tr,
  paragraphPos,
  paragraph,
  next,
  nextContent,
  styleResolver,
}: JoinAtParagraphMarkOptions): void => {
  const boundary = paragraphPos + paragraph.nodeSize;
  const join = paragraphJoinAfter(tr.doc, boundary);
  if (!join || join.next !== next)
    panic("A paragraph-mark join lost its next paragraph");
  if (join.inlineBookmarks.length > 0) {
    const leftToken = getProseParagraphPropertySourceToken(paragraph);
    const rightToken = getProseParagraphPropertySourceToken(next);
    tr.replaceWith(boundary - 1, join.nextPosition + 1, join.inlineBookmarks);
    tr.setNodeMarkup(
      paragraphPos,
      undefined,
      joinedParagraphAttrs({
        paragraph,
        next,
        between: join.inlineBookmarks.reduce(
          (size, node) => size + node.nodeSize,
          0,
        ),
        nextContent,
      }),
    );
    markParagraphPropertySourceTransfers(tr, [
      {
        displacedToken: typeof leftToken === "string" ? leftToken : null,
        selectedToken: typeof rightToken === "string" ? rightToken : null,
      },
    ]);
  } else {
    joinProseParagraphsWithRightPropertySource({
      attrs: joinedParagraphAttrs({ paragraph, next, nextContent }),
      pos: boundary,
      transaction: tr,
    });
  }
  if (!styleResolver) return;
  if (paragraph.content.size > 0) {
    rebaseParagraphRuns({
      previousContext: paragraphRunStyleContext(paragraph, styleResolver),
      paragraphPosition: paragraphPos,
      range: { from: 0, to: paragraph.content.size },
      styleResolver,
      tr,
    });
  }
  tr.setMeta(JOINED_RUNS_RESTYLED_META, true);
};
