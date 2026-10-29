/**
 * The document fonts folio bundles: one table for every place that names them.
 *
 * `@stll/folio-fonts` ships the files, and its generator
 * (`scripts/generate-folio-fonts.ts`) reads this table to decide which files
 * to copy and which `@font-face` rules to write, so the published package and
 * the resolver here cannot disagree about a family. The module imports nothing:
 * the generator, the resolver and the editor's availability check all read it
 * without pulling in the rest of the engine.
 */

const PACKAGE_TABLE = {
  Arimo: "arimo",
  Caladea: "caladea",
  Carlito: "carlito",
  Cousine: "cousine",
  Lato: "lato",
  "Noto Sans Arabic": "noto-sans-arabic",
  "Source Sans 3": "source-sans-3",
  Tinos: "tinos",
} as const;

export type FontsourcePackageTable = typeof PACKAGE_TABLE;

/** Each bundled family and the directory (an `@fontsource` package name) that ships it. */
export const FONTSOURCE_PACKAGES: FontsourcePackageTable = PACKAGE_TABLE;

export type BundledFamily = keyof FontsourcePackageTable;

/**
 * What each bundled family stands in for, as prose for the package's notices.
 * The resolver's own mapping lives in `utils/fontLoader.ts` and
 * `utils/fontResolver.ts`.
 */
export const BUNDLED_FAMILY_SUBSTITUTES: Readonly<
  Record<BundledFamily, string>
> = {
  Arimo: "Arial, and the generic sans-serif",
  Caladea: "Cambria",
  Carlito: "Calibri",
  Cousine: "Courier New, and the generic monospace",
  Lato: "Aptos and Aptos Display",
  "Noto Sans Arabic": "Dubai, and Arabic script in every other family",
  "Source Sans 3": "Segoe UI",
  Tinos: "Times New Roman, and the generic serif",
};

/**
 * The subsets a headless measurer or the PDF writer reads, in resolution
 * priority: a code point two subsets cover resolves to the earlier one.
 * A browser is served every subset the family is cut into.
 */
export const MEASURED_SUBSETS: readonly ["latin", "latin-ext", "arabic"] = [
  "latin",
  "latin-ext",
  "arabic",
];

export type MeasuredSubset = (typeof MEASURED_SUBSETS)[number];

/** Folio paints 400 and 700 only, as `DisplayFontFace` states. */
export const BUNDLED_WEIGHTS: readonly [400, 700] = [400, 700];

export type BundledWeight = (typeof BUNDLED_WEIGHTS)[number];

/** The styles a family may be cut in; a family without italics ships `normal` only. */
export const BUNDLED_STYLES: readonly ["normal", "italic"] = [
  "normal",
  "italic",
];

export type BundledStyle = (typeof BUNDLED_STYLES)[number];

/** `files/<package>-<subset>-<weight>-<style>.<format>`: the path of one served binary. */
export const bundledFontFilePath = ({
  packageName,
  subset,
  weight,
  style,
  format,
}: {
  readonly packageName: string;
  readonly subset: string;
  readonly weight: number;
  readonly style: BundledStyle;
  readonly format: "woff" | "woff2";
}): string =>
  `files/${packageName}-${subset}-${String(weight)}-${style}.${format}`;

/** Each family's directory carries the ranges its subsets were cut with. */
export const UNICODE_RANGE_FILE = "unicode.json";
