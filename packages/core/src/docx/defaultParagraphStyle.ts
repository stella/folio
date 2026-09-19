/**
 * Which style a document means when it says "the default paragraph style".
 *
 * ECMA-376 17.7.4.17 (`w:default`): the default for a style type is the style
 * that carries it, and where several styles of one type do, the last one wins.
 * A document that marks none has no document-defined default and the consumer
 * applies its own built-in, which a document redefines under the built-in
 * *name* rather than under a fixed style id: `w:styleId` is opaque, so a
 * localized or generated package names its default `Standard`, `Normln` or
 * `style0` while still writing `<w:name w:val="Normal"/>`. Matching the id
 * alone finds English Word output and nothing else.
 */

import type { Style } from "../types/document";

/** The name Word's built-in default paragraph style always carries in the file. */
export const BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME = "Normal";

/** The id English Word gives it, kept as the last resort for the reverse case. */
export const BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID = "Normal";

/**
 * The formatting Word's default template gives Normal: 8pt (160 twips) after
 * spacing, 1.08x line spacing.
 *
 * A package that declares neither a default paragraph style nor `w:docDefaults`
 * renders as though it declared this, so anything that mints the missing
 * default has to mint the formatting with it.
 */
export const BUILT_IN_DEFAULT_PARAGRAPH_FORMATTING = {
  spaceAfter: 160,
  lineSpacing: 259,
  lineSpacingRule: "auto",
} satisfies NonNullable<Style["pPr"]>;

type MintDefaultParagraphStyleOptions = {
  takenStyleIds: ReadonlySet<string>;
  /** Whether the source declared `w:docDefaults`, which the set carries over. */
  hasDocDefaults: boolean;
};

/**
 * The default paragraph style a set needs when its source declared none.
 *
 * The id only has to be free, because the set is what defines it. The
 * formatting has to be the built-in template's whenever the source had no
 * `w:docDefaults`, because that is what the source itself rendered as: a
 * consumer applies its built-in Normal only where no default paragraph style
 * exists, and this minted style is one. Where the source did declare
 * `w:docDefaults`, the set carries them and they remain authoritative, so the
 * minted style states nothing.
 */
export const mintDefaultParagraphStyle = ({
  takenStyleIds,
  hasDocDefaults,
}: MintDefaultParagraphStyleOptions): Style => {
  let styleId = BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID;
  for (let suffix = 1; takenStyleIds.has(styleId); suffix += 1) {
    styleId = `${BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID}${suffix}`;
  }
  return {
    styleId,
    type: "paragraph",
    name: BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME,
    default: true,
    ...(hasDocDefaults ? {} : { pPr: { ...BUILT_IN_DEFAULT_PARAGRAPH_FORMATTING } }),
  };
};

export const resolveDefaultParagraphStyle = (styles: Iterable<Style>): Style | undefined => {
  let flagged: Style | undefined;
  let namedBuiltIn: Style | undefined;
  let idBuiltIn: Style | undefined;
  for (const style of styles) {
    if (style.type !== "paragraph") {
      continue;
    }
    if (style.default) {
      flagged = style;
    }
    if (namedBuiltIn === undefined && style.name === BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME) {
      namedBuiltIn = style;
    }
    if (idBuiltIn === undefined && style.styleId === BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID) {
      idBuiltIn = style;
    }
  }
  return flagged ?? namedBuiltIn ?? idBuiltIn;
};
