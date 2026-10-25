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
import { RtlExtension } from "./marks/RtlExtension";
import { RunFormattingOverrideExtension } from "./marks/RunFormattingOverrideExtension";
import { RunIdentityExtension } from "./marks/RunIdentityExtension";
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

export const MARK_EXTENSIONS: {
  readonly bold: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly italic: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly underline: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly strike: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly textColor: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly runShading: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly highlight: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly fontSize: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly fontFamily: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly language: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly superscript: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly subscript: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly hyperlink: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly allCaps: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly smallCaps: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly footnoteRef: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly characterSpacing: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly emboss: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly imprint: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly hidden: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly textShadow: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly emphasisMark: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly textOutline: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly rtl: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly textEffect: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly runFormattingOverride: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly characterStyle: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly runIdentity: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly comment: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly insertion: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly deletion: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly runPropertyChange: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
  readonly inlineWrapper: (
    options?: Partial<Record<string, unknown>> | undefined,
  ) => MarkExtension;
} = {
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
  runIdentity: RunIdentityExtension,
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
 * Five of these marks are containers on the save leg as well, and their
 * relative order is that leg's rather than a choice made here.
 * `extractParagraphContent` opens a comment range around everything a leaf
 * produces. For an editor-created revision, the existing wrapper stack owns
 * the change, followed by the hyperlink and the runs it holds. So the save
 * writes `w:commentRangeStart … w:bdo > w:ins > w:hyperlink > w:r`, and the
 * editor DOM nests `span.docx-comment > bdo > span.docx-insertion > a`.
 * `markNestingOrder.test.ts` holds the two legs to each other.
 *
 * `runIdentity` is the fifth: `w:r` is the element it names, and it is the
 * innermost element that leg emits.
 *
 * Every other mark is a run property or run content: it has no element of its
 * own in OOXML, so the save leg ranks it against nothing and its place here is
 * a presentation decision.
 */
export const MARK_NESTING_ORDER: readonly [
  "bold",
  "italic",
  "underline",
  "strike",
  "textColor",
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
  "comment",
  "inlineWrapper",
  "insertion",
  "deletion",
  "hyperlink",
  "runPropertyChange",
  "runIdentity",
] = markNestingOrder([
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
  // The container tier, in the save leg's order.
  "comment",
  // An editor-created revision belongs inside the wrapper stack already on
  // its content, so an inserted override saves as
  // `w:bdo > w:ins > w:hyperlink > w:r`.
  "inlineWrapper",
  "insertion",
  "deletion",
  "hyperlink",
  "runPropertyChange",
  // Last, because `w:r` is the innermost element the save leg emits: every
  // container above ranks against it, and it ranks against nothing.
  "runIdentity",
]);
