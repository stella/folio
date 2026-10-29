/**
 * The bundled faces' bytes for Node.js: a headless measurer, a PDF export, a
 * server-rendered preview.
 *
 * ```ts
 * import { createFontsourceFaces } from "@stll/folio-core/fonts/fontsourceFaces";
 * import { folioFontFiles } from "@stll/folio-fonts/node";
 *
 * const faces = createFontsourceFaces(folioFontFiles);
 * ```
 *
 * Files are found beside this module, so the package works from any
 * `node_modules` layout. A bundle that moves the code away from the files
 * passes their directory to {@link createFolioFontFiles}.
 */

import { readFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { FOLIO_FONT_FILES, folioFontPath } from "./index";
import type { FolioFontFiles } from "./types";

export type { FolioFontFiles } from "./types";

// Both `src/node.ts` and the built `dist/node.js` sit one level below the
// package root. The VS Code extension's bundle rewrites this exact line.
const FONTS_ROOT = new URL("../fonts/", import.meta.url);

/** The absolute path of the package's `fonts/` directory. */
export const folioFontsDirectory = (): string =>
  resolve(fileURLToPath(FONTS_ROOT));

const SHIPPED: ReadonlySet<string> = new Set(FOLIO_FONT_FILES);

const isMissing = (error: unknown): boolean =>
  error instanceof Error &&
  "code" in error &&
  (error.code === "ENOENT" || error.code === "ENOTDIR");

/**
 * A reader over a `fonts/` directory laid out as this package ships it. Only
 * a file the package ships is read: any other name, including one that would
 * leave the directory, answers `null`.
 */
export const createFolioFontFiles = (
  directory: string | URL = FONTS_ROOT,
): FolioFontFiles => {
  const root =
    typeof directory === "string" ? directory : fileURLToPath(directory);
  return {
    read: (packageName, relativePath) => {
      const file = folioFontPath(packageName, relativePath);
      if (!SHIPPED.has(file)) return null;
      try {
        return new Uint8Array(
          readFileSync(join(root, packageName, relativePath)),
        );
      } catch (error) {
        if (isMissing(error)) return null;
        throw error;
      }
    },
  };
};

/** The package's own files. */
export const folioFontFiles: FolioFontFiles = createFolioFontFiles();
