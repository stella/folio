/**
 * The repository's bundled faces for the paint harness: the shared
 * `createFontsourceFaces` table, reading the files `@stll/folio-fonts` keeps
 * under `packages/fonts/fonts`. Measurement, the PDF writer, and the browser's `@font-face` rules
 * resolve each code point to the same binary.
 */

import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  createFontsourceFaces,
  type FontsourceFontFaceCssOptions,
} from "../src/fonts/fontsourceFaces";
import type { HeadlessFontSource } from "../src/fonts/headlessMeasure";

export { cssString } from "../src/fonts/fontsourceFaces";

/** `<root>/packages/core/scripts`, resolved from this file, never from the cwd. */
const REPO_ROOT = path.resolve(import.meta.dir, "..", "..", "..");

const FONTS_DIR = path.join(REPO_ROOT, "packages", "fonts", "fonts");

const faces = createFontsourceFaces({
  read: (packageName, relativePath) => {
    const filePath = path.join(FONTS_DIR, packageName, relativePath);
    return existsSync(filePath) ? new Uint8Array(readFileSync(filePath)) : null;
  },
});

export const createBundledFontSource = (): HeadlessFontSource => faces.source;

export type BundledFontFaceCssOptions = FontsourceFontFaceCssOptions;

export const bundledFontFaceCss = (
  options: BundledFontFaceCssOptions = {},
): string => faces.fontFaceCss(options);
