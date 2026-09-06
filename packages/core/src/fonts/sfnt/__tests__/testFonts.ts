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

/**
 * `@fontsource` ships a family cut into disjoint script subsets: `latin` has
 * ASCII but no `Ř`, `latin-ext` has `Ř` but no `A`. A Czech, Slovak or Polish
 * document needs both at once, so a fixture has to be able to name one.
 */
export const TEST_FONT_SUBSETS = ["latin", "latin-ext"] as const;

export type TestFontSubset = (typeof TEST_FONT_SUBSETS)[number];

type TestFontFaceOptions = {
  readonly family: TestFontFamily;
  readonly weight?: 400 | 700;
  readonly style?: "normal" | "italic";
  readonly subset?: TestFontSubset;
};

/** Path of one `@fontsource` WOFF face, resolved from this module's location. */
export const testFontPath = ({
  family,
  weight = 400,
  style = "normal",
  subset = "latin",
}: TestFontFaceOptions): string =>
  join(
    import.meta.dir,
    "../../../../../react/node_modules/@fontsource",
    family,
    "files",
    `${family}-${subset}-${weight}-${style}.woff`,
  );

/**
 * Whether every fixture face is on disk. Checked once, synchronously, so the
 * suites can decide to skip before any test body runs.
 */
export const TEST_FONTS_INSTALLED = TEST_FONT_FAMILIES.every((family) =>
  TEST_FONT_SUBSETS.every((subset) => existsSync(testFontPath({ family, subset }))),
);

/** Reason shown on skipped suites, so an absent fixture is never a mystery. */
export const TEST_FONTS_SKIP_REASON =
  "needs the @fontsource fixtures installed under packages/react/node_modules";

export const readTestFont = async (options: TestFontFaceOptions): Promise<Uint8Array> =>
  new Uint8Array(await Bun.file(testFontPath(options)).arrayBuffer());
