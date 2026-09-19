/**
 * Every style a package points at must be a style that package defines.
 *
 * A `w:pStyle`, `w:rStyle` or `w:tblStyle` naming an id `styles.xml` never
 * declares is a dangling reference. Word tolerates one by falling back to the
 * default style, which is exactly why it goes unnoticed: the document looks
 * almost right, the paragraph silently loses its formatting, and a consumer
 * that classifies by style — folio's own outline, TOC and AI snapshot included
 * — sees nothing there. `w:basedOn`, `w:next` and `w:link` inside `styles.xml`
 * dangle the same way.
 *
 * This is the "dangling by construction" bug class: a writer that hardcodes an
 * English built-in id it does not also define produces it for *every* document
 * it writes, not for some unlucky input. {@link danglingStyleReferences} is the
 * shared check, so an authoring path can be held to it by a test rather than by
 * the reviewer noticing.
 */

const STYLE_REFERENCE = /<w:(?:pStyle|rStyle|tblStyle)\b[^>]*\bw:val="(?<id>[^"]*)"/gu;
const STYLE_LINK = /<w:(?:basedOn|next|link)\b[^>]*\bw:val="(?<id>[^"]*)"/gu;
const STYLE_DEFINITION = /<w:style\b[^>]*\bw:styleId="(?<id>[^"]*)"/gu;

/** One reference that resolves to nothing, and the part it was written into. */
export type DanglingStyleReference = {
  /** The `w:val` that names no defined style. */
  styleId: string;
  /** The package part the reference sits in, e.g. `word/document.xml`. */
  part: string;
};

const idsIn = (xml: string, pattern: RegExp): string[] => {
  const ids: string[] = [];
  for (const match of xml.matchAll(pattern)) {
    const id = match.groups?.["id"];
    if (id !== undefined && id.length > 0) {
      ids.push(id);
    }
  }
  return ids;
};

/**
 * Collect the style references in `parts` that `stylesXml` does not define.
 *
 * `parts` maps a part name to its XML; pass every part that can carry a style
 * reference (`word/document.xml`, the header/footer parts, `word/footnotes.xml`,
 * `word/endnotes.xml`, `word/comments.xml`, and `word/styles.xml` itself for
 * the `basedOn`/`next`/`link` graph).
 */
export const danglingStyleReferences = (
  stylesXml: string | undefined,
  parts: ReadonlyMap<string, string>,
): DanglingStyleReference[] => {
  const defined = new Set(idsIn(stylesXml ?? "", STYLE_DEFINITION));
  const dangling: DanglingStyleReference[] = [];
  for (const [part, xml] of parts) {
    const references =
      part === "word/styles.xml"
        ? idsIn(xml, STYLE_LINK)
        : [...idsIn(xml, STYLE_REFERENCE), ...idsIn(xml, STYLE_LINK)];
    for (const styleId of references) {
      if (!defined.has(styleId)) {
        dangling.push({ styleId, part });
      }
    }
  }
  return dangling;
};
