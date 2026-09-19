/**
 * Whether a paragraph is a heading, and of which level, without reading the
 * style id.
 *
 * A style id is opaque. ECMA-376 17.7.4.17 says a style id "can be assigned in
 * any manner desired" and only requires it to be unique within the document,
 * and Word fills it by stripping spaces and non-ASCII characters out of the
 * *localized* UI name. The same built-in heading therefore arrives as
 * `Heading1`, `berschrift1` (de), `Nadpis1` (cs), `Titre1` (fr), `Nagwek1`
 * (pl), `Cmsor1` (hu), `Ttulo1` (es), `Overskrift1` (da) or a bare `1`.
 * Matching the id finds English Word output and nothing else, the same trap
 * `defaultParagraphStyle.ts` documents for `Normal`.
 *
 * Two signals survive localisation, in this order:
 *
 * 1. The effective `w:outlineLvl` — direct paragraph formatting, else the
 *    style chain (`styleParser` flattens `w:basedOn` into `style.pPr` per
 *    17.7.1). 17.3.1.20 gives it the range 0..9 and says it "shall be used to
 *    calculate the TOC field", which is what a `TOC \u` field and the
 *    navigation pane read. So a custom style with an outline level *is* a
 *    heading, and a heading-named style reset to
 *    {@link BODY_TEXT_OUTLINE_LEVEL} is *not* (folio's own `TOCHeading`, based
 *    on `Heading1`, does exactly that).
 * 2. The style's `w:name`. Part 1 has no normative list of built-in names, but
 *    it keys latent styles by name alone: a `w:lsdException` carries no style
 *    id, and 17.7.4.8 defines its `w:name` as "the primary name for the style
 *    which shall inherit this set of latent style property exceptions",
 *    ignored when "the current application does not know of an internal
 *    primary style with the current name". Annex L writes those names in
 *    English (`heading 1`…`heading 9`) with no locale qualification.
 *
 * The id is never a signal on its own.
 *
 * Measured over the 5,341-file public corpus: 207 files define a heading whose
 * id is not `HeadingN` while the name is English, covering 123 distinct ids;
 * only 4 documents localise the *name*, all ODF→OOXML exports (LibreOffice
 * writing its ODF display names verbatim). Names are not case-stable (5.2% are
 * `Heading 1` rather than the Annex L `heading 1`) and not space-stable
 * (LibreOffice 25.8 writes `Heading1`, and `IntenseQuote` for `Intense
 * Quote`), which is why {@link normalizeStyleName} ignores both.
 */

import type { DocDefaults, Style } from "../types/document";
import {
  BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME,
  resolveDefaultParagraphStyle,
} from "./defaultParagraphStyle";

/**
 * The tenth `w:outlineLvl` value. 17.3.1.20: "the val attribute … can be from
 * 0 to 9, where 9 specifically indicates that there is no outline level
 * specifically applied to this paragraph." It is a deliberate "not a heading",
 * not a tenth level. Every range test goes through
 * {@link isHeadingOutlineLevel} so the reserved value keeps one meaning across
 * the codebase.
 *
 * The same clause adds that an omitted element "is assumed to be 9". That
 * default cannot be applied to a *style* definition, because 17.7.1 tells
 * producers not to write a property "already been set by a previous level of
 * the style hierarchy": a document that names a style `heading 1` and omits
 * the level is inheriting the consumer's built-in definition, which carries
 * level 0. An absent level therefore means "unspecified, ask the name", and
 * only a written 9 means body text.
 */
export const BODY_TEXT_OUTLINE_LEVEL = 9;

/** The highest `w:outlineLvl` that still names a heading (outline level nine). */
const MAX_HEADING_OUTLINE_LEVEL = 8;

/** True when an outline level names a heading rather than body text. */
export const isHeadingOutlineLevel = (level: number | null | undefined): level is number =>
  typeof level === "number" &&
  Number.isInteger(level) &&
  level >= 0 &&
  level <= MAX_HEADING_OUTLINE_LEVEL;

/**
 * Compare style names the way producers actually write them. The corpus shows
 * both `heading 1` (Annex L, 94.8%) and `Heading 1` (5.2%), and LibreOffice
 * drops the space entirely (`Heading1`, `IntenseQuote`), so case and
 * whitespace are the tolerance. A name is otherwise matched whole: a style a
 * Czech template calls `Nadpis 1` stays a custom style.
 */
export const normalizeStyleName = (name: string): string =>
  name.trim().toLowerCase().replace(/\s+/gu, "");

/**
 * The `w:name` Word itself writes for each built-in, and therefore the
 * spelling every style table folio authors must use. One owner: a style set and
 * the classifier that reads it cannot drift apart if both name the same
 * constant.
 *
 * Word is not uniformly cased and guessing gets it wrong, so each value is the
 * spelling that dominates Microsoft Word output in the public corpus:
 * `footnote text` 375 against 49 `Footnote Text`, `footer` 1145 against 2,
 * `caption` 390 against 60 — but `Body Text` 424 against 2, `Title` 621
 * against 2, and the auto-generated linked character styles (`Footnote Text
 * Char` 248, `Endnote Text Char` 116) title-cased without exception.
 * {@link normalizeStyleName} makes matching tolerant of all of it; this map is
 * about what folio *writes*.
 */
export const BUILT_IN_STYLE_NAME = {
  /** 4,515 Word occurrences against 3 lowercase. */
  normal: BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME,
  bodyText: "Body Text",
  title: "Title",
  subtitle: "Subtitle",
  quote: "Quote",
  intenseQuote: "Intense Quote",
  listParagraph: "List Paragraph",
  tocHeading: "TOC Heading",
  caption: "caption",
  header: "header",
  footer: "footer",
  footnoteText: "footnote text",
  commentReference: "annotation reference",
  footnoteReference: "footnote reference",
  footnoteTextChar: "Footnote Text Char",
  endnoteText: "endnote text",
  endnoteReference: "endnote reference",
  endnoteTextChar: "Endnote Text Char",
  hyperlink: "Hyperlink",
  defaultParagraphFont: "Default Paragraph Font",
  noList: "No List",
  normalTable: "Normal Table",
  tableGrid: "Table Grid",
} as const;

type BuiltInStyleName = (typeof BUILT_IN_STYLE_NAME)[keyof typeof BUILT_IN_STYLE_NAME];

/**
 * The name of a built-in heading, from its zero-based outline level.
 * Lowercase: 1,066 Word occurrences of `heading 1` against 13 `Heading 1`, and
 * Annex L writes the latent-style exceptions the same way.
 */
export const builtInHeadingStyleName = (outlineLevel: number): string =>
  `heading ${outlineLevel + 1}`;

/** The name of a built-in TOC entry style, from its one-based level (`toc 1`). */
export const builtInTableOfContentsStyleName = (level: number): string => `toc ${level}`;

/**
 * `heading 1`…`heading 9` against an already-normalised name. Word's built-in
 * names stop at nine, matching `w:outlineLvl`'s nine heading values, so
 * `Cmsor10` → `Címsor 10` is a user style rather than a tenth built-in.
 */
const BUILT_IN_HEADING_NAME = /^heading(?<level>[1-9])$/u;

/** `toc 1`…`toc 9`, the styles a `TOC` field writes its entries in. */
const BUILT_IN_TABLE_OF_CONTENTS_NAME = /^toc(?<level>[1-9])$/u;

/**
 * The outline level a built-in heading *name* implies (zero-based, so
 * `heading 1` is 0), or undefined when the name is not a built-in heading.
 */
const headingOutlineLevelFromStyleName = (name: string | undefined): number | undefined => {
  if (name === undefined) {
    return undefined;
  }
  const level = BUILT_IN_HEADING_NAME.exec(normalizeStyleName(name))?.groups?.["level"];
  return level === undefined ? undefined : Number.parseInt(level, 10) - 1;
};

/**
 * A document's styles indexed by what they *are* rather than by what they are
 * called. Built once per document: the consumers below classify every
 * paragraph, and rebuilding the maps per paragraph would make each of them
 * quadratic.
 */
export type BuiltInStyleIndex = {
  /** The style's effective `w:outlineLvl`, including 9, or undefined. */
  outlineLevelOf: (styleId: string | null | undefined) => number | undefined;
  /** The zero-based level a built-in heading *name* implies, or undefined. */
  headingLevelFromNameOf: (styleId: string | null | undefined) => number | undefined;
  /** The built-in this style is, by name, or undefined for a custom style. */
  builtInNameOf: (styleId: string | null | undefined) => BuiltInStyleName | undefined;
  /**
   * The level an English built-in heading *id* implies, and only when the
   * package defines no style under it. See {@link resolveHeadingLevel} tier 3.
   */
  undefinedBuiltInHeadingLevelOf: (styleId: string | null | undefined) => number | undefined;
  /** The document's style id for a built-in heading level (zero-based). */
  styleIdForHeadingLevel: (level: number) => string | undefined;
  /** The document's style id for a built-in TOC entry level (one-based, `toc 1`). */
  styleIdForTableOfContentsLevel: (level: number) => string | undefined;
  /** The document's style id for a named built-in, e.g. `TOC Heading`. */
  styleIdForBuiltInName: (name: BuiltInStyleName) => string | undefined;
};

/**
 * Every spelling a producer might write, mapped back to the canonical one, so
 * a caller compares against {@link BUILT_IN_STYLE_NAME} rather than against a
 * normalised form it would have to spell a second time.
 */
const CANONICAL_BY_NORMALIZED = new Map<string, BuiltInStyleName>(
  Object.values(BUILT_IN_STYLE_NAME).map((name) => [normalizeStyleName(name), name]),
);

const asBuiltInName = (normalized: string): BuiltInStyleName | undefined =>
  CANONICAL_BY_NORMALIZED.get(normalized);

/**
 * The outline level a style chain sets: the style's own `w:outlineLvl`, else
 * the nearest ancestor's (17.7.1). `styleParser` already flattens `w:basedOn`
 * for a parsed package, but a style set built in memory carries the raw chain,
 * so the walk keeps both kinds of input on the same answer. `seen` guards the
 * circular `basedOn` a malformed package can contain.
 */
const inheritedOutlineLevel = (
  style: Style,
  styleById: ReadonlyMap<string, Style>,
): number | undefined => {
  const seen = new Set<string>();
  let current: Style | undefined = style;
  while (current && !seen.has(current.styleId)) {
    seen.add(current.styleId);
    if (current.pPr?.outlineLevel !== undefined) {
      return current.pPr.outlineLevel;
    }
    current = current.basedOn === undefined ? undefined : styleById.get(current.basedOn);
  }
  return undefined;
};

export const createBuiltInStyleIndex = (
  styles: Iterable<Style>,
  docDefaults?: DocDefaults | undefined,
): BuiltInStyleIndex => {
  const styleIdByHeadingLevel = new Map<number, string>();
  const styleIdByTableOfContentsLevel = new Map<number, string>();
  const styleIdByBuiltInName = new Map<BuiltInStyleName, string>();

  const paragraphStyles: Style[] = [];
  const paragraphStyleById = new Map<string, Style>();
  const styleById = new Map<string, Style>();
  for (const style of styles) {
    // `w:basedOn` may point at any style type, so the chain walk needs them
    // all even though only paragraph styles are classified.
    styleById.set(style.styleId, style);
    if (style.type === "paragraph") {
      paragraphStyles.push(style);
      paragraphStyleById.set(style.styleId, style);
    }
  }

  for (const style of paragraphStyles) {
    if (style.name === undefined) {
      continue;
    }
    const headingLevel = headingOutlineLevelFromStyleName(style.name);
    // First definition wins, matching `resolveDefaultParagraphStyle`: a package
    // that names two styles `heading 1` is malformed, and taking the first
    // keeps the choice deterministic.
    if (headingLevel !== undefined && !styleIdByHeadingLevel.has(headingLevel)) {
      styleIdByHeadingLevel.set(headingLevel, style.styleId);
      continue;
    }
    const normalized = normalizeStyleName(style.name);
    const tocLevel = BUILT_IN_TABLE_OF_CONTENTS_NAME.exec(normalized)?.groups?.["level"];
    if (tocLevel !== undefined) {
      const level = Number.parseInt(tocLevel, 10);
      if (!styleIdByTableOfContentsLevel.has(level)) {
        styleIdByTableOfContentsLevel.set(level, style.styleId);
      }
      continue;
    }
    const builtInName = asBuiltInName(normalized);
    if (builtInName !== undefined && !styleIdByBuiltInName.has(builtInName)) {
      styleIdByBuiltInName.set(builtInName, style.styleId);
    }
  }

  /**
   * The style a paragraph actually resolves against. ECMA-376 17.7.2 layer 3
   * is "the paragraph's own style chain", and a paragraph with no `w:pStyle`
   * (or one naming a style the package never defines, or one naming a
   * character style) takes the default paragraph style — the same fallback
   * `StyleResolver` applies, so the model and the editor cannot drift apart.
   */
  const defaultStyle = resolveDefaultParagraphStyle(paragraphStyles);
  const styleFor = (styleId: string | null | undefined): Style | undefined =>
    (styleId === null || styleId === undefined ? undefined : paragraphStyleById.get(styleId)) ??
    defaultStyle;

  // Layer 1 of the same cascade. The schema allows `w:outlineLvl` in
  // `w:docDefaults/w:pPrDefault`, and the style chain overrides it.
  const docDefaultOutlineLevel = docDefaults?.pPr?.outlineLevel;

  const outlineLevelCache = new Map<string | null | undefined, number | undefined>();

  return {
    outlineLevelOf: (styleId) => {
      if (outlineLevelCache.has(styleId)) {
        return outlineLevelCache.get(styleId);
      }
      const style = styleFor(styleId);
      const level =
        (style === undefined ? undefined : inheritedOutlineLevel(style, styleById)) ??
        docDefaultOutlineLevel;
      outlineLevelCache.set(styleId, level);
      return level;
    },
    headingLevelFromNameOf: (styleId) => headingOutlineLevelFromStyleName(styleFor(styleId)?.name),
    builtInNameOf: (styleId) => {
      const name = styleFor(styleId)?.name;
      return name === undefined ? undefined : asBuiltInName(normalizeStyleName(name));
    },
    undefinedBuiltInHeadingLevelOf: (styleId) =>
      styleId === null || styleId === undefined || paragraphStyleById.has(styleId)
        ? undefined
        : headingOutlineLevelFromStyleName(styleId),
    styleIdForHeadingLevel: (level) => styleIdByHeadingLevel.get(level),
    styleIdForTableOfContentsLevel: (level) => styleIdByTableOfContentsLevel.get(level),
    styleIdForBuiltInName: (name) => styleIdByBuiltInName.get(name),
  };
};

/** An index over a document that defines no styles: every lookup misses. */
export const EMPTY_BUILT_IN_STYLE_INDEX: BuiltInStyleIndex = createBuiltInStyleIndex([]);

/**
 * What a consumer knows about a paragraph: its style id and whatever
 * `w:outlineLvl` applies to it. The ProseMirror `outlineLevel` attr already
 * holds direct-else-style resolution, so passing it here agrees with passing
 * direct formatting from the DOCX model.
 */
export type ParagraphOutlineSource = {
  styleId?: string | null | undefined;
  outlineLevel?: number | null | undefined;
};

/**
 * The heading level a paragraph carries, zero-based (`heading 1` is 0), or
 * undefined when it is not a heading.
 *
 * Precedence:
 *
 * 1. An effective outline level decides on its own, including
 *    {@link BODY_TEXT_OUTLINE_LEVEL}, which means "not a heading". A style
 *    named `heading 5` whose outline level is 0 is a level-1 heading; a style
 *    named `heading 3` reset to 9 is body text. The format gives the outline
 *    level to field calculation (17.3.1.20) and leaves the name to the UI
 *    (17.7.4.9), so the level is the one the document asserts.
 * 2. Only when no outline level is set anywhere does the built-in `w:name`
 *    decide — the style is then inheriting the consumer's own built-in
 *    definition, which supplies the level.
 * 3. Last resort, and only for a `w:pStyle` the package defines no style for:
 *    the id itself, read as the English built-in id. 17.7.4.17 makes a style
 *    without `w:customStyle` a built-in and lets an application recognise it
 *    "if the associated style ID is known", which is the one case where the id
 *    is all the information left. `defaultParagraphStyle.ts` keeps the same
 *    last tier for `Normal`. A document that defines its heading styles never
 *    reaches this, so it cannot override a name or an outline level — and a
 *    localized package never writes an English id to begin with.
 *
 * Two consequences of rule 1 are deliberate, not oversights.
 *
 * **An outline level on a style that is not a heading still makes a heading.**
 * The corpus has 680 such occurrences across 161 files, including `Title` at
 * level 0 (42×) and `Subtitle` at level 1 (23×), plus `H1`, `Sub-heading`,
 * `index heading` and a `DSTOC1-1`…`DSTOC8-8` family. Setting the level is how
 * a document asks for a paragraph to be outlined, and Word's navigation pane
 * and a `TOC \u` field both honour it, so folio does not second-guess a
 * document that asked. Suppressing `Title` here would mean folio deciding a
 * document's outline differs from Word's.
 *
 * **An outline level that disagrees with a built-in heading name wins.** 26
 * corpus styles do this (`heading 5` at level 0, `heading 3` at level 1, and
 * so on), all from non-Word producers or hand-authored fixtures. The format
 * gives the level to field calculation (17.3.1.20) and the name to the user
 * interface (17.7.4.9), so the level is the machine-readable claim and the
 * name is a label. This is the one rule below that was not confirmed against
 * Word itself.
 */
export const resolveHeadingLevel = (
  paragraph: ParagraphOutlineSource,
  index: BuiltInStyleIndex,
): number | undefined => {
  const effective = paragraph.outlineLevel ?? index.outlineLevelOf(paragraph.styleId);
  if (effective !== null && effective !== undefined) {
    return isHeadingOutlineLevel(effective) ? effective : undefined;
  }
  return (
    index.headingLevelFromNameOf(paragraph.styleId) ??
    index.undefinedBuiltInHeadingLevelOf(paragraph.styleId)
  );
};

/** True when the paragraph's style is Word's `Quote` or `Intense Quote`. */
export const isQuoteStyle = (
  styleId: string | null | undefined,
  index: BuiltInStyleIndex,
): boolean => {
  const name = index.builtInNameOf(styleId);
  return name === BUILT_IN_STYLE_NAME.quote || name === BUILT_IN_STYLE_NAME.intenseQuote;
};
