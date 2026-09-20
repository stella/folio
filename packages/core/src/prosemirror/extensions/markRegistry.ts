/**
 * Every mark the editor schema declares, and the order they nest in.
 *
 * `StarterKit` builds the schema from this record and from nothing else, so
 * `SchemaMarkName` is the set of marks that exist rather than a list beside it
 * — a table that must decide something per mark (what the painter does with
 * it) can be checked total against this union, and a mark added without that
 * decision does not compile.
 *
 * The record answers which marks exist and how each is built; its key order
 * answers nothing. `MARK_NESTING_ORDER` is the registration order, and
 * registration order is DOM nesting order.
 */

import { AllCapsExtension } from "./marks/AllCapsExtension";
import { BoldExtension } from "./marks/BoldExtension";
import { CharacterSpacingExtension } from "./marks/CharacterSpacingExtension";
import { CharacterStyleExtension } from "./marks/CharacterStyleExtension";
import { CommentExtension } from "./marks/CommentExtension";
import { FontFamilyExtension } from "./marks/FontFamilyExtension";
import { FontSizeExtension } from "./marks/FontSizeExtension";
import { FootnoteRefExtension } from "./marks/FootnoteRefExtension";
import { HiddenTextExtension } from "./marks/HiddenTextExtension";
import { HighlightExtension } from "./marks/HighlightExtension";
import { HyperlinkExtension } from "./marks/HyperlinkExtension";
import { InlineWrapperExtension } from "./marks/InlineWrapperExtension";
import { ItalicExtension } from "./marks/ItalicExtension";
import { LanguageExtension } from "./marks/LanguageExtension";
import { PageBreakRunOwnerExtension } from "./marks/PageBreakRunOwnerExtension";
import { RtlExtension } from "./marks/RtlExtension";
import { RunFormattingOverrideExtension } from "./marks/RunFormattingOverrideExtension";
import { RunShadingExtension } from "./marks/RunShadingExtension";
import { SmallCapsExtension } from "./marks/SmallCapsExtension";
import { StrikeExtension } from "./marks/StrikeExtension";
import { SubscriptExtension } from "./marks/SubscriptExtension";
import { SuperscriptExtension } from "./marks/SuperscriptExtension";
import { TextColorExtension } from "./marks/TextColorExtension";
import { TextEffectExtension } from "./marks/TextEffectExtension";
import {
  EmbossExtension,
  EmphasisMarkExtension,
  ImprintExtension,
  TextOutlineExtension,
  TextShadowExtension,
} from "./marks/TextEffectsExtensions";
import {
  DeletionExtension,
  InsertionExtension,
  RunPropertyChangeExtension,
} from "./marks/TrackedChangeExtensions";
import { UnderlineExtension } from "./marks/UnderlineExtension";
import type { MarkExtension } from "./types";

export const MARK_EXTENSIONS = {
  bold: BoldExtension,
  italic: ItalicExtension,
  underline: UnderlineExtension,
  strike: StrikeExtension,
  textColor: TextColorExtension,
  runShading: RunShadingExtension,
  highlight: HighlightExtension,
  fontSize: FontSizeExtension,
  fontFamily: FontFamilyExtension,
  language: LanguageExtension,
  superscript: SuperscriptExtension,
  subscript: SubscriptExtension,
  hyperlink: HyperlinkExtension,
  allCaps: AllCapsExtension,
  smallCaps: SmallCapsExtension,
  footnoteRef: FootnoteRefExtension,
  characterSpacing: CharacterSpacingExtension,
  emboss: EmbossExtension,
  imprint: ImprintExtension,
  hidden: HiddenTextExtension,
  textShadow: TextShadowExtension,
  emphasisMark: EmphasisMarkExtension,
  textOutline: TextOutlineExtension,
  rtl: RtlExtension,
  textEffect: TextEffectExtension,
  runFormattingOverride: RunFormattingOverrideExtension,
  characterStyle: CharacterStyleExtension,
  pageBreakRunOwner: PageBreakRunOwnerExtension,
  comment: CommentExtension,
  insertion: InsertionExtension,
  deletion: DeletionExtension,
  runPropertyChange: RunPropertyChangeExtension,
  inlineWrapper: InlineWrapperExtension,
} as const satisfies Record<string, () => MarkExtension>;

/** A mark the editor schema declares. */
export type SchemaMarkName = keyof typeof MARK_EXTENSIONS;

/** The order, rejected unless it names every mark; a repeat is the test's. */
const markNestingOrder = <const Names extends readonly SchemaMarkName[]>(
  names: Names & (SchemaMarkName extends Names[number] ? unknown : never),
): Names => names;

/**
 * Outermost first: the order the editor registers marks in, which is the order
 * it nests their DOM elements in.
 *
 * Four of these marks are containers on the save leg as well, and their
 * relative order is that leg's rather than a choice made here.
 * `extractParagraphContent` opens a comment range around everything a leaf
 * produces (it flushes the open run and hyperlink and drops the tracked wrapper
 * before pushing the marker: `fromProseDoc.ts:2359-2369`), then the revision
 * wrapper (`fromProseDoc.ts:2530`), then the wrapper stack inside that revision
 * (`nestInlineWrapperGroups`, `fromProseDoc.ts:2200`), then the hyperlink and
 * the runs it holds (`fromProseDoc.ts:2567`). So the save writes
 * `w:commentRangeStart … w:ins > w:bdo > w:hyperlink > w:r`, and the editor DOM
 * nests `span.docx-comment > span.docx-insertion > bdo > a`.
 * `markNestingOrder.test.ts` holds the two legs to each other.
 *
 * Every other mark is a run property or run content: it has no element of its
 * own in OOXML, so the save leg ranks it against nothing and its place here is
 * a presentation decision.
 */
export const MARK_NESTING_ORDER = markNestingOrder([
  "bold",
  "italic",
  "underline",
  "strike",
  "textColor",
  // Before highlight: a later mark is the inner DOM span, so an explicit
  // highlight wins the background when a run has both — matching the paged
  // painter's `run.highlight ?? run.shading`. (#722)
  "runShading",
  "highlight",
  "fontSize",
  "fontFamily",
  "language",
  "superscript",
  "subscript",
  "allCaps",
  "smallCaps",
  "footnoteRef",
  "characterSpacing",
  "emboss",
  "imprint",
  "hidden",
  "textShadow",
  "emphasisMark",
  "textOutline",
  "rtl",
  "textEffect",
  "runFormattingOverride",
  "characterStyle",
  "pageBreakRunOwner",
  // The container tier, in the save leg's order.
  "comment",
  "insertion",
  "deletion",
  // Inside the revision and outside the link: a wrapper group is cut before
  // the run, the hyperlink and the revision the walk had open
  // (`fromProseDoc.ts:2396-2405`), and a revision in the group keeps its place
  // and takes the nest inside it (`fromProseDoc.ts:2196-2202`), so an inserted
  // override saves as `w:ins > w:bdo > w:hyperlink > w:r`. (#912)
  "inlineWrapper",
  "hyperlink",
  "runPropertyChange",
]);
