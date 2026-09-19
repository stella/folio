import { describe, expect, test } from "bun:test";

import {
  BODY_TEXT_OUTLINE_LEVEL,
  createBuiltInStyleIndex,
  isQuoteStyle,
  resolveHeadingLevel,
} from "./builtInStyles";
import type { Style } from "../types/document";

const paragraphStyle = (style: Omit<Style, "type">): Style => ({ ...style, type: "paragraph" });

/** What a localized Word writes: the UI name stripped of spaces and accents. */
const LOCALIZED_HEADINGS = [
  { styleId: "Nadpis1", name: "heading 1", level: 0, locale: "cs" },
  { styleId: "berschrift2", name: "heading 2", level: 1, locale: "de" },
  { styleId: "Titre3", name: "heading 3", level: 2, locale: "fr" },
  { styleId: "Nagwek4", name: "heading 4", level: 3, locale: "pl" },
  { styleId: "Cmsor5", name: "heading 5", level: 4, locale: "hu" },
  { styleId: "Ttulo6", name: "heading 6", level: 5, locale: "es" },
  { styleId: "Overskrift7", name: "heading 7", level: 6, locale: "da" },
  // Word's remap for an unknown style leaves a bare ordinal as the id.
  { styleId: "8", name: "heading 8", level: 7, locale: "opaque" },
  // A generated package names its styles by position.
  { styleId: "style7", name: "heading 9", level: 8, locale: "generated" },
] as const;

describe("heading classification by built-in name", () => {
  for (const { styleId, name, level, locale } of LOCALIZED_HEADINGS) {
    test(`${styleId} (${locale}) named "${name}" is heading level ${level + 1}`, () => {
      const index = createBuiltInStyleIndex([paragraphStyle({ styleId, name })]);
      expect(resolveHeadingLevel({ styleId }, index)).toBe(level);
    });
  }

  test("a non-ASCII style id classifies by its name", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Überschrift1", name: "heading 1" }),
      paragraphStyle({ styleId: "标题2", name: "heading 2" }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Überschrift1" }, index)).toBe(0);
    expect(resolveHeadingLevel({ styleId: "标题2" }, index)).toBe(1);
  });

  test("the name is matched case- and space-insensitively", () => {
    // 5.2% of the public corpus writes `Heading 1`, and LibreOffice 25.8
    // writes `Heading1` with no space at all.
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "a", name: "Heading 1" }),
      paragraphStyle({ styleId: "b", name: "Heading2" }),
      paragraphStyle({ styleId: "c", name: "HEADING  3" }),
    ]);
    expect(resolveHeadingLevel({ styleId: "a" }, index)).toBe(0);
    expect(resolveHeadingLevel({ styleId: "b" }, index)).toBe(1);
    expect(resolveHeadingLevel({ styleId: "c" }, index)).toBe(2);
  });

  test("a heading-looking name in another language is not a built-in", () => {
    // The localized name is what the UI shows; only the file-format name
    // counts. These come from ODF exports, which write display names verbatim.
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "style1", name: "Überschrift 1" }),
      paragraphStyle({ styleId: "style2", name: "Nadpis 2" }),
      paragraphStyle({ styleId: "style3", name: "Nagłówek 3" }),
      paragraphStyle({ styleId: "style4", name: "Címsor 4" }),
      paragraphStyle({ styleId: "style5", name: "Encabezado 5" }),
    ]);
    for (const styleId of ["style1", "style2", "style3", "style4", "style5"]) {
      expect(resolveHeadingLevel({ styleId }, index)).toBeUndefined();
    }
  });

  test("an English style id alone is not a heading when the package defines it", () => {
    // The package defines `Heading1` and says nothing else about it: no
    // built-in name, no outline level. The id is not evidence.
    const index = createBuiltInStyleIndex([paragraphStyle({ styleId: "Heading1" })]);
    expect(resolveHeadingLevel({ styleId: "Heading1" }, index)).toBeUndefined();
  });

  test("a dangling English built-in id is read as the built-in it names", () => {
    // A minimal `styles.xml` that defines only `Normal` while the body styles
    // paragraphs `Heading1`/`Heading2`: Word applies its own built-in
    // definition, so the id is the only surviving signal. A localized package
    // never writes these ids, so this tier cannot mislocalise anything.
    const index = createBuiltInStyleIndex([paragraphStyle({ styleId: "Normal", name: "Normal" })]);
    expect(resolveHeadingLevel({ styleId: "Heading1" }, index)).toBe(0);
    expect(resolveHeadingLevel({ styleId: "heading2" }, index)).toBe(1);
    expect(resolveHeadingLevel({ styleId: "Nadpis1" }, index)).toBeUndefined();
    expect(resolveHeadingLevel({ styleId: "ClauseHeading1" }, index)).toBeUndefined();
  });

  test("a defined style's own answer beats the dangling-id reading", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Heading1", name: "Body Text", pPr: { outlineLevel: 9 } }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Heading1" }, index)).toBeUndefined();
  });

  test("a heading name above nine is a user style", () => {
    const index = createBuiltInStyleIndex([paragraphStyle({ styleId: "x", name: "heading 10" })]);
    expect(resolveHeadingLevel({ styleId: "x" }, index)).toBeUndefined();
  });
});

describe("heading classification by outline level", () => {
  test("a custom style with only an outline level is a heading", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Clause", name: "Clause Heading", pPr: { outlineLevel: 1 } }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Clause" }, index)).toBe(1);
  });

  test("a direct outline level makes a Normal paragraph a heading", () => {
    const index = createBuiltInStyleIndex([paragraphStyle({ styleId: "Normal", name: "Normal" })]);
    expect(resolveHeadingLevel({ styleId: "Normal", outlineLevel: 2 }, index)).toBe(2);
  });

  test("outline level nine is body text, not a tenth level", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Body", name: "Body Text", pPr: { outlineLevel: 9 } }),
    ]);
    expect(BODY_TEXT_OUTLINE_LEVEL).toBe(9);
    expect(resolveHeadingLevel({ styleId: "Body" }, index)).toBeUndefined();
    expect(resolveHeadingLevel({ styleId: "Body", outlineLevel: 9 }, index)).toBeUndefined();
  });

  test("an outline level of nine beats a built-in heading name", () => {
    // `TOC Heading` is based on `heading 1` and resets the level: it titles the
    // table of contents, it is not an entry in it.
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Nadpis1", name: "heading 1", pPr: { outlineLevel: 0 } }),
      paragraphStyle({
        styleId: "Obsah",
        name: "TOC Heading",
        basedOn: "Nadpis1",
        pPr: { outlineLevel: 9 },
      }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Obsah" }, index)).toBeUndefined();
  });

  test("the outline level wins when it disagrees with the name", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "H5", name: "heading 5", pPr: { outlineLevel: 0 } }),
    ]);
    expect(resolveHeadingLevel({ styleId: "H5" }, index)).toBe(0);
  });

  test("direct formatting wins over the style's level", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "H1", name: "heading 1", pPr: { outlineLevel: 0 } }),
    ]);
    expect(resolveHeadingLevel({ styleId: "H1", outlineLevel: 9 }, index)).toBeUndefined();
    expect(resolveHeadingLevel({ styleId: "H1", outlineLevel: 3 }, index)).toBe(3);
  });

  test("an outline level is inherited through basedOn", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Base", name: "Custom Base", pPr: { outlineLevel: 1 } }),
      paragraphStyle({ styleId: "Derived", name: "Custom Derived", basedOn: "Base" }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Derived" }, index)).toBe(1);
  });

  test("a circular basedOn chain terminates", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "A", name: "A", basedOn: "B" }),
      paragraphStyle({ styleId: "B", name: "B", basedOn: "A" }),
    ]);
    expect(resolveHeadingLevel({ styleId: "A" }, index)).toBeUndefined();
  });

  test("a paragraph with no style takes the default paragraph style", () => {
    // Pathological but legal: the document's default paragraph style carries
    // an outline level, so every unstyled paragraph is a heading. The editor's
    // cascade resolves it that way, so the classifier must too.
    const index = createBuiltInStyleIndex([
      paragraphStyle({
        styleId: "Standard",
        name: "Normal",
        default: true,
        pPr: { outlineLevel: 1 },
      }),
    ]);
    expect(resolveHeadingLevel({ styleId: undefined }, index)).toBe(1);
    // An id the package never defines falls back the same way.
    expect(resolveHeadingLevel({ styleId: "Heading1" }, index)).toBe(1);
  });

  test("an outline level in docDefaults applies when no style sets one", () => {
    const index = createBuiltInStyleIndex([paragraphStyle({ styleId: "Normal", name: "Normal" })], {
      pPr: { outlineLevel: 2 },
    });
    expect(resolveHeadingLevel({ styleId: "Normal" }, index)).toBe(2);
    // The style chain overrides it.
    const overridden = createBuiltInStyleIndex(
      [paragraphStyle({ styleId: "Normal", name: "Normal", pPr: { outlineLevel: 9 } })],
      { pPr: { outlineLevel: 2 } },
    );
    expect(resolveHeadingLevel({ styleId: "Normal" }, overridden)).toBeUndefined();
  });

  test("an unknown or absent style id is not a heading", () => {
    const index = createBuiltInStyleIndex([]);
    expect(resolveHeadingLevel({ styleId: "Nadpis1" }, index)).toBeUndefined();
    expect(resolveHeadingLevel({ styleId: undefined }, index)).toBeUndefined();
    expect(resolveHeadingLevel({ styleId: null, outlineLevel: null }, index)).toBeUndefined();
  });
});

describe("other built-in styles", () => {
  test("Quote and Intense Quote resolve through localized ids", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Zitat", name: "Quote" }),
      paragraphStyle({ styleId: "Citationintense", name: "Intense Quote" }),
      // LibreOffice writes the name without the space.
      paragraphStyle({ styleId: "Idzet", name: "IntenseQuote" }),
      paragraphStyle({ styleId: "BlockQuoteCustom", name: "Block Quote Custom" }),
    ]);
    expect(isQuoteStyle("Zitat", index)).toBe(true);
    expect(isQuoteStyle("Citationintense", index)).toBe(true);
    expect(isQuoteStyle("Idzet", index)).toBe(true);
    expect(isQuoteStyle("BlockQuoteCustom", index)).toBe(false);
    expect(isQuoteStyle(undefined, index)).toBe(false);
  });

  test("Title and Subtitle are built-ins but not headings by name", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "Cm", name: "Title" }),
      paragraphStyle({ styleId: "Alcm", name: "Subtitle" }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Cm" }, index)).toBeUndefined();
    expect(resolveHeadingLevel({ styleId: "Alcm" }, index)).toBeUndefined();
  });

  test("the document's own id for a built-in is discoverable", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({ styleId: "berschrift1", name: "heading 1" }),
      paragraphStyle({ styleId: "berschrift2", name: "heading 2" }),
    ]);
    expect(index.styleIdForHeadingLevel(0)).toBe("berschrift1");
    expect(index.styleIdForHeadingLevel(1)).toBe("berschrift2");
    expect(index.styleIdForHeadingLevel(2)).toBeUndefined();
  });

  test("a heading style is indexed under the level it resolves to", () => {
    // The name says level 5, the outline level says level 1, and
    // `resolveHeadingLevel` believes the outline level. The index has to agree
    // with it, or a caller asking for level 5 gets a style Word outlines at 1.
    const style = paragraphStyle({
      styleId: "Heading5",
      name: "heading 5",
      pPr: { outlineLevel: 0 },
    });
    const index = createBuiltInStyleIndex([style]);
    expect(resolveHeadingLevel({ styleId: "Heading5" }, index)).toBe(0);
    expect(index.styleIdForHeadingLevel(0)).toBe("Heading5");
    expect(index.styleIdForHeadingLevel(4)).toBeUndefined();
  });

  test("a heading style reset to body text is indexed at no level", () => {
    const index = createBuiltInStyleIndex([
      paragraphStyle({
        styleId: "Heading3",
        name: "heading 3",
        pPr: { outlineLevel: BODY_TEXT_OUTLINE_LEVEL },
      }),
    ]);
    expect(resolveHeadingLevel({ styleId: "Heading3" }, index)).toBeUndefined();
    expect(index.styleIdForHeadingLevel(2)).toBeUndefined();
  });

  test("a character style named like a heading is ignored", () => {
    const index = createBuiltInStyleIndex([
      { styleId: "Heading1Char", type: "character", name: "heading 1" },
    ]);
    expect(resolveHeadingLevel({ styleId: "Heading1Char" }, index)).toBeUndefined();
    expect(index.styleIdForHeadingLevel(0)).toBeUndefined();
  });
});
