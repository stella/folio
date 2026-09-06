/**
 * Test support: real font bytes borrowed from the `@fontsource` packages that
 * `@stll/folio-react` installs.
 *
 * No font binary is copied into this repository. A checkout without those
 * packages installed simply has no fixtures, and the suites that need them
 * skip themselves rather than failing.
 */

import { existsSync } from "node:fs";
import { join } from "node:path";

export const TEST_FONT_FAMILIES = ["arimo", "caladea", "carlito", "cousine", "tinos"] as const;

export type TestFontFamily = (typeof TEST_FONT_FAMILIES)[number];

type TestFontFaceOptions = {
  readonly family: TestFontFamily;
  readonly weight?: 400 | 700;
  readonly style?: "normal" | "italic";
};

/** Path of one `@fontsource` WOFF face, resolved from this module's location. */
export const testFontPath = ({
  family,
  weight = 400,
  style = "normal",
}: TestFontFaceOptions): string =>
  join(
    import.meta.dir,
    "../../../../../react/node_modules/@fontsource",
    family,
    "files",
    `${family}-latin-${weight}-${style}.woff`,
  );

/**
 * Whether every fixture face is on disk. Checked once, synchronously, so the
 * suites can decide to skip before any test body runs.
 */
export const TEST_FONTS_INSTALLED = TEST_FONT_FAMILIES.every((family) =>
  existsSync(testFontPath({ family })),
);

/** Reason shown on skipped suites, so an absent fixture is never a mystery. */
export const TEST_FONTS_SKIP_REASON =
  "needs the @fontsource fixtures installed under packages/react/node_modules";

export const readTestFont = async (options: TestFontFaceOptions): Promise<Uint8Array> =>
  new Uint8Array(await Bun.file(testFontPath(options)).arrayBuffer());
