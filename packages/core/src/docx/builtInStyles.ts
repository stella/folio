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
import { resolveDefaultParagraphStyle } from "./defaultParagraphStyle";

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
export const MAX_HEADING_OUTLINE_LEVEL = 8;

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

/** Normalised `w:name` values this module recognises by name alone. */
export const BUILT_IN_STYLE_NAMES = {
  title: "title",
  subtitle: "subtitle",
  quote: "quote",
  intenseQuote: "intensequote",
  listParagraph: "listparagraph",
  tocHeading: "tocheading",
} as const;

export type BuiltInStyleName = (typeof BUILT_IN_STYLE_NAMES)[keyof typeof BUILT_IN_STYLE_NAMES];

/**
 * `heading 1`…`heading 9` against an already-normalised name. Word's built-in
 * names stop at nine, matching `w:outlineLvl`'s nine heading values, so
 * `Cmsor10` → `Címsor 10` is a user style rather than a tenth built-in.
 */
const BUILT_IN_HEADING_NAME = /^heading(?<level>[1-9])$/u;

/**
 * The outline level a built-in heading *name* implies (zero-based, so
 * `heading 1` is 0), or undefined when the name is not a built-in heading.
 */
export const headingOutlineLevelFromStyleName = (name: string | undefined): number | undefined => {
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
  /** The document's style id for a built-in name, or undefined when it defines none. */
  styleIdForBuiltInName: (name: BuiltInStyleName) => string | undefined;
  /** The document's style id for a built-in heading level (zero-based). */
  styleIdForHeadingLevel: (level: number) => string | undefined;
};

const BUILT_IN_NAMES: ReadonlySet<string> = new Set(Object.values(BUILT_IN_STYLE_NAMES));

const asBuiltInName = (normalized: string): BuiltInStyleName | undefined =>
  // SAFETY: `BUILT_IN_NAMES` holds exactly the `BuiltInStyleName` values, so a
  // hit narrows the string to one of them.
  BUILT_IN_NAMES.has(normalized) ? (normalized as BuiltInStyleName) : undefined;

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
  const styleIdByBuiltInName = new Map<BuiltInStyleName, string>();
  const styleIdByHeadingLevel = new Map<number, string>();

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
    if (headingLevel !== undefined) {
      // First definition wins, matching `resolveDefaultParagraphStyle`: a
      // package that names two styles `heading 1` is malformed, and taking the
      // first keeps the choice deterministic.
      if (!styleIdByHeadingLevel.has(headingLevel)) {
        styleIdByHeadingLevel.set(headingLevel, style.styleId);
      }
      continue;
    }
    const builtInName = asBuiltInName(normalizeStyleName(style.name));
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
    styleIdForBuiltInName: (name) => styleIdByBuiltInName.get(name),
    styleIdForHeadingLevel: (level) => styleIdByHeadingLevel.get(level),
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
 * Note that rule 1 makes a `Title` or `Subtitle` carrying an outline level a
 * heading. That is the document's own assertion and what Word's outline reads;
 * this module does not second-guess it.
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
  return name === BUILT_IN_STYLE_NAMES.quote || name === BUILT_IN_STYLE_NAMES.intenseQuote;
};
