/**
 * The document fonts folio measures and paints with, as data: which families
 * ship, which faces each is cut into, and where each file sits in the package.
 *
 * Nothing here reads a file, so it runs anywhere. Import
 * `@stll/folio-fonts/fonts.css` to declare the faces in a browser, a single
 * file as `@stll/folio-fonts/fonts/<path>` to hand it to a bundler, or
 * `@stll/folio-fonts/node` to read the bytes in Node.js.
 */

import {
  FACES,
  FAMILIES,
  type FolioFontFamilyName,
  PACKAGES,
} from "./catalogue.gen";
import type { FolioFontFace, FolioFontFamily } from "./types";

export type { FolioFontFamilyName } from "./catalogue.gen";
export type { FolioFontFace, FolioFontFamily, FolioFontFiles } from "./types";

/** Each bundled family and its directory under `fonts/`. */
export const FOLIO_FONT_PACKAGES: Readonly<
  Record<FolioFontFamilyName, string>
> = PACKAGES;

/** The bundled families, with their licence, source and role. */
export const FOLIO_FONT_FAMILIES: readonly FolioFontFamily[] = FAMILIES;

/**
 * Every served face, one entry per subset, in the order `fonts.css` declares
 * them.
 */
export const FOLIO_FONT_FACES: readonly FolioFontFace[] = FACES;

/** The package-relative path of the stylesheet that declares every face. */
export const FOLIO_FONTS_STYLESHEET = "fonts.css";

/** The files each family directory carries besides its faces. */
const FAMILY_FILES: readonly string[] = ["unicode.json", "LICENSE"];

/**
 * The package-relative path of a family's file, as `FolioFontFiles.read`
 * names it: `folioFontPath("arimo", "unicode.json")` is
 * `fonts/arimo/unicode.json`.
 */
export const folioFontPath = (
  packageName: string,
  relativePath: string,
): string => `fonts/${packageName}/${relativePath}`;

/** Every file under `fonts/`, package-relative. */
export const FOLIO_FONT_FILES: readonly string[] = [
  ...Object.values(PACKAGES).flatMap((packageName) =>
    FAMILY_FILES.map((file) => folioFontPath(packageName, file)),
  ),
  ...FACES.flatMap(({ woff2, woff }) =>
    woff === null ? [woff2] : [woff2, woff],
  ),
];
