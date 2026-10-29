/** A bundled family: where it comes from and what it stands in for. */
export type FolioFontFamily = {
  /** The CSS family name the faces are declared under, such as `Carlito`. */
  readonly family: string;
  /** Its directory under `fonts/`, named after its `@fontsource` package. */
  readonly packageName: string;
  /** The SPDX identifier the font files are licensed under. */
  readonly license: string;
  /** The `@fontsource` release the files were copied from. */
  readonly fontsourceVersion: string;
  /** The authored family names it stands in for, as prose. */
  readonly substitutes: string;
  /** Where the family is developed. */
  readonly source: string;
};

/** One served binary: a weight and style of a family, cut to one subset. */
export type FolioFontFace = {
  readonly family: string;
  readonly packageName: string;
  /** The `@fontsource` subset, such as `latin` or `latin-ext`. */
  readonly subset: string;
  readonly weight: 400 | 700;
  readonly style: "normal" | "italic";
  /** The CSS `unicode-range` the subset is declared with. */
  readonly unicodeRange: string;
  /** Package-relative path of the `.woff2` a browser loads. */
  readonly woff2: string;
  /**
   * Package-relative path of the `.woff` a headless measurer or PDF writer
   * parses, or `null` for a subset folio does not measure with.
   */
  readonly woff: string | null;
};

/**
 * Reads a family's files by the path `@fontsource` lays them out under, the
 * shape `createFontsourceFaces` in `@stll/folio-core` reads through.
 */
export type FolioFontFiles = {
  /** Bytes of `fonts/<packageName>/<relativePath>`, or `null` when there is no such file. */
  readonly read: (
    packageName: string,
    relativePath: string,
  ) => Uint8Array | null;
};
