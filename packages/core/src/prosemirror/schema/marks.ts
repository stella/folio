/**
 * ProseMirror Mark Type Interfaces
 *
 * Type definitions for mark attributes used by conversion modules,
 * extensions, and other consumers. MarkSpec definitions have moved
 * to the extension system (extensions/marks/).
 */

import type { ShadingProperties } from "../../types/colors";
import type {
  EmphasisMark,
  RunPropertyChange,
  TextEffect,
  TextFormatting,
  ThemeColorSlot,
  UnderlineStyle,
} from "../../types/document";

/**
 * Text color mark attributes
 */
export type TextColorAttrs = {
  rgb?: string;
  themeColor?: ThemeColorSlot;
  themeTint?: string;
  themeShade?: string;
};

/**
 * Underline mark attributes
 */
export type UnderlineAttrs = {
  style?: UnderlineStyle;
  color?: TextColorAttrs;
};

export type StrikeAttrs = {
  double?: boolean;
};

/**
 * Font size mark attributes
 */
export type FontSizeAttrs = {
  size: number; // in half-points (OOXML format)
};

/**
 * Font family mark attributes
 */
export type FontFamilyAttrs = {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  hint?: NonNullable<NonNullable<TextFormatting["fontFamily"]>["hint"]>;
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
};

export type LanguageAttrs = {
  val?: string;
  eastAsia?: string;
  bidi?: string;
};

export type HighlightAttrs = {
  color: NonNullable<TextFormatting["highlight"]>;
};

/**
 * Run-level shading mark attributes (w:shd). Carries the shading FILL as a
 * flattened ColorValue (mirroring TextColorAttrs) plus the pattern. The default
 * `clear` pattern is never stored (absence ⇒ the fill renders as a solid
 * background and re-serializes to `w:val="clear"`); only non-clear patterns
 * (`pct*`, stripes) are carried for export fidelity (`solid` is flattened into
 * the fill). `patternColor` is the pattern foreground rgb (`w:shd w:color`),
 * carried so a non-clear pattern's color survives export.
 */
export type RunShadingAttrs = TextColorAttrs & {
  pattern?: NonNullable<ShadingProperties["pattern"]>;
  patternColor?: string;
};

export type CharacterSpacingAttrs = {
  spacing?: number;
  position?: number;
  scale?: number;
  kerning?: number;
};

export type EmphasisMarkAttrs = {
  type?: Exclude<EmphasisMark, "none">;
};

/**
 * Text effect mark attributes (w:effect). The "none" sentinel is never marked;
 * absence of the mark is the no-effect state.
 */
export type TextEffectAttrs = {
  effect: Exclude<TextEffect, "none">;
};

export type FootnoteRefAttrs = {
  id: string | number;
  noteType?: "footnote" | "endnote";
  vertAlign?: "baseline" | "superscript";
};

export type CommentAttrs = {
  commentId: number;
};

/**
 * Provenance of a tracked change.
 *
 * - `"user"` — a normal tracked change authored by a person (or an
 *   accepted suggestion). Serializes to OOXML `w:ins` / `w:del` /
 *   `w:rPrChange`.
 * - `"suggested"` — an AI-proposed edit rendered with the tracked-change
 *   visual grammar but ALWAYS stripped from serialized DOCX output until
 *   accepted. Suggested marks never round-trip through OOXML; parsing a
 *   DOCX can only ever produce `"user"` marks.
 */
export const TRACKED_CHANGE_PROVENANCE_VALUES = ["user", "suggested"] as const;
export type TrackedChangeProvenance = (typeof TRACKED_CHANGE_PROVENANCE_VALUES)[number];

export type TrackedChangeMarkAttrs = {
  revisionId: number;
  author: string;
  date?: string;
  /** UTC companion timestamp carried privately through the editable model. */
  utcDate?: string;
  /** Optional author initials (w:initials) carried through the round-trip. */
  initials?: string;
  moveKind?: "moveTo" | "moveFrom";
  /** Defaults to `"user"`; `"suggested"` for AI-proposed, non-serialized edits. */
  provenance: TrackedChangeProvenance;
  /**
   * Groups every mark belonging to one logical suggestion so the host can
   * accept/reject them together. Absent for `"user"` provenance (the mark's
   * `null` default is normalized to absent by the attrs reader).
   */
  suggestionId?: string;
};

/** Run-property revisions carried through the editable model. */
export type RunPropertyChangeMarkAttrs = {
  changes: RunPropertyChange[];
  /** See {@link TrackedChangeProvenance}. Defaults to `"user"`. */
  provenance: TrackedChangeProvenance;
  suggestionId?: string;
};

/** Editor-only identity for one authored run that contains an explicit page break. */
export type PageBreakRunOwnerMarkAttrs = {
  id: number;
};

export const COMPLEX_SCRIPT_RUN_PROPERTY_KEYS = ["boldCs", "italicCs", "fontSizeCs"] as const;

export type ComplexScriptRunPropertyKey = (typeof COMPLEX_SCRIPT_RUN_PROPERTY_KEYS)[number];

export const RUN_FORMATTING_BOOLEAN_PROPERTIES = [
  "bold",
  "boldCs",
  "italic",
  "italicCs",
  "strike",
  "doubleStrike",
  "smallCaps",
  "allCaps",
  "hidden",
  "emboss",
  "imprint",
  "outline",
  "shadow",
  "rtl",
  "cs",
] as const;

export const RUN_FORMATTING_VALUE_PROPERTIES = [
  "underline",
  "vertAlign",
  "color",
  "highlight",
  "shading",
  "fontSize",
  "fontSizeCs",
  "fontFamily",
  "language",
  "spacing",
  "position",
  "scale",
  "kerning",
  "effect",
  "emphasisMark",
] as const;

export type RunFormattingBooleanProperty = (typeof RUN_FORMATTING_BOOLEAN_PROPERTIES)[number];
export type RunFormattingValueProperty = (typeof RUN_FORMATTING_VALUE_PROPERTIES)[number];

type RunFormattingPropertySpec =
  Extract<RunFormattingBooleanProperty, RunFormattingValueProperty> extends never
    ? {
        [Property in keyof TextFormatting]: Property extends RunFormattingBooleanProperty
          ? "boolean"
          : Property extends RunFormattingValueProperty
            ? "value"
            : Property extends "styleId"
              ? "style"
              : never;
      }
    : never;

export const RUN_FORMATTING_PROPERTY_SPECS = {
  bold: "boolean",
  boldCs: "boolean",
  italic: "boolean",
  italicCs: "boolean",
  underline: "value",
  strike: "boolean",
  doubleStrike: "boolean",
  vertAlign: "value",
  smallCaps: "boolean",
  allCaps: "boolean",
  hidden: "boolean",
  color: "value",
  highlight: "value",
  shading: "value",
  fontSize: "value",
  fontSizeCs: "value",
  fontFamily: "value",
  language: "value",
  spacing: "value",
  position: "value",
  scale: "value",
  kerning: "value",
  effect: "value",
  emphasisMark: "value",
  emboss: "boolean",
  imprint: "boolean",
  outline: "boolean",
  shadow: "boolean",
  rtl: "boolean",
  cs: "boolean",
  styleId: "style",
} as const satisfies RunFormattingPropertySpec;

export type AuthoredRunFormattingValues = Partial<Pick<TextFormatting, RunFormattingValueProperty>>;

export type RunFormattingOverrideAttrs = {
  allCaps?: boolean;
  bold?: boolean;
  boldCs?: boolean;
  cs?: boolean;
  emboss?: boolean;
  hidden?: boolean;
  imprint?: boolean;
  italic?: boolean;
  italicCs?: boolean;
  outline?: boolean;
  shadow?: boolean;
  smallCaps?: boolean;
  strike?: boolean;
  /** Imported/reconciled direct-positive baseline; current PM signals may diverge after edits. */
  _authoredOn?: readonly RunFormattingBooleanProperty[];
  /** Imported/reconciled direct-negative baseline; absence means the property was inherited. */
  _authoredOff?: readonly RunFormattingBooleanProperty[];
  /** Imported/reconciled direct value baseline, including nested slot identity and sentinels. */
  _authoredValues?: AuthoredRunFormattingValues;
  directFontProperties?: readonly ("fontFamily" | "fontSize" | "color")[];
  /** Current complex-script mirrors explicitly omitted by the run. */
  complexScriptPropertyAbsences?: readonly ComplexScriptRunPropertyKey[];
  color?: "auto";
  doubleStrike?: false;
  effect?: "none";
  emphasisMark?: "none";
  highlight?: "none";
  kerning?: 0;
  position?: 0;
  rtl?: false;
  scale?: 100;
  shading?: ShadingProperties & { pattern: "nil" };
  spacing?: 0;
  fontSizeCs?: number;
  underline?: "none";
  vertAlign?: "baseline";
};

/**
 * Character style mark attributes (w:rStyle).
 *
 * `styleId` is the OOXML character style reference, carried so a styled run
 * re-serializes as a style reference instead of losing the semantic link.
 * Style formatting is resolved once per document through the style engine;
 * duplicating it on every styled run bloats collaborative state and becomes
 * stale when a run moves between paragraph contexts.
 */
export type CharacterStyleAttrs = {
  styleId: string;
};

/**
 * Hyperlink mark attributes
 */
export type HyperlinkAttrs = {
  href: string;
  tooltip?: string;
  rId?: string;
  _docxHyperlinkIndex?: number;
};
