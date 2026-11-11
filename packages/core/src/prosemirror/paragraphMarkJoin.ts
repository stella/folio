/**
 * Removing a paragraph mark: the one join both a resolved revision and a
 * retracted one perform.
 *
 * A paragraph's mark ends it, and in OOXML its properties live on that mark.
 * With the mark gone, the paragraph runs on into the next one and ends with
 * the NEXT paragraph's mark, so the paragraph left is the next one: its
 * identity, its paragraph properties (style, numbering, the mark's own run
 * properties), its mark's revision and its section endpoint all survive, and
 * the removed mark takes its paragraph's properties and section endpoint with
 * it. This holds whether the first paragraph still has words or not.
 *
 * The first paragraph's runs keep their direct formatting; what its style
 * lent them is re-read in the next paragraph's style.
 */

import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

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

/** The two paragraphs a join runs together, and what stands between them. */
type JoinSides = {
  /** The paragraph whose mark goes. */
  paragraph: PMNode;
  /** The paragraph after it, whose mark stays. */
  next: PMNode;
  /** Inline size placed between the two paragraphs' content (bookmarks). */
  between: number;
};

/**
 * Where one attribute of the joined paragraph comes from: `"next"` for the
 * mark that stays (its properties, its revision, its identity), or a rule
 * that reads both sides.
 */
type JoinedAttrRule = "next" | ((sides: JoinSides) => unknown);

const nextValue = (key: string) => (sides: JoinSides) => sides.next.attrs[key];

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
  // Identity: the paragraph left is the next one.
  paraId: "next",
  textId: "next",
  idStability: "next",
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

/**
 * The attrs of the paragraph a join leaves, each by its rule in
 * `JOINED_PARAGRAPH_ATTRS`. `between` is the inline size the join places
 * between the two paragraphs' content.
 */
export const joinedParagraphAttrs = (
  paragraph: PMNode,
  next: PMNode,
  between = 0,
): PMNode["attrs"] => {
  const sides = { paragraph, next, between };
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
      joinedParagraphAttrs(
        paragraph,
        next,
        join.inlineBookmarks.reduce((size, node) => size + node.nodeSize, 0),
      ),
    );
    markParagraphPropertySourceTransfers(tr, [
      {
        displacedToken: typeof leftToken === "string" ? leftToken : null,
        selectedToken: typeof rightToken === "string" ? rightToken : null,
      },
    ]);
  } else {
    joinProseParagraphsWithRightPropertySource({
      attrs: joinedParagraphAttrs(paragraph, next),
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
