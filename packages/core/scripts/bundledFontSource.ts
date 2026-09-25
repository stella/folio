/**
 * The repository's `@fontsource` faces for the paint harness: the shared
 * `createFontsourceFaces` table, reading the packages the React adapter
 * installs. Measurement, the PDF writer, and the browser's `@font-face` rules
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

const FONTSOURCE_DIR = path.join(REPO_ROOT, "packages", "react", "node_modules", "@fontsource");

const faces = createFontsourceFaces({
  read: (packageName, relativePath) => {
    const filePath = path.join(FONTSOURCE_DIR, packageName, relativePath);
    return existsSync(filePath) ? new Uint8Array(readFileSync(filePath)) : null;
  },
});

export const createBundledFontSource = (): HeadlessFontSource => faces.source;

export type BundledFontFaceCssOptions = FontsourceFontFaceCssOptions;

export const bundledFontFaceCss = (options: BundledFontFaceCssOptions = {}): string =>
  faces.fontFaceCss(options);
