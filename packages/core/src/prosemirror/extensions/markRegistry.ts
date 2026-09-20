/**
 * Every mark the editor schema declares, in registration order.
 *
 * Registration order is DOM nesting order: a mark registered later is the
 * inner span. `StarterKit` builds the schema from this record and from nothing
 * else, so `SchemaMarkName` is the set of marks that exist rather than a list
 * beside it — a table that must decide something per mark (what the painter
 * does with it) can be checked total against this union, and a mark added
 * without that decision does not compile.
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
  // Registered BEFORE highlight: a later mark is the inner DOM span, so an
  // explicit highlight (registered after) wins the background when a run has
  // both — matching the paged painter's `run.highlight ?? run.shading`. (#722)
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
  // Registered last: the wrapper is what a revision's own span sits outside
  // of, so an inserted override reads `<span class="docx-insertion"><bdo …>`.
  inlineWrapper: InlineWrapperExtension,
} as const satisfies Record<string, () => MarkExtension>;

/** A mark the editor schema declares. */
export type SchemaMarkName = keyof typeof MARK_EXTENSIONS;
