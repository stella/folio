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

/**
 * Faces for the scripts that shape. Kept apart from the Latin fixtures because
 * they come from this package's own devDependencies rather than the adapter's
 * runtime ones: a consumer of `@stll/folio-react` has no use for a Devanagari
 * face, and a test of shaping cannot do without one.
 */
export const SHAPING_TEST_FACES = {
  arabic:
    "../../../../../react/node_modules/@fontsource/noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff",
  devanagari:
    "../../../../node_modules/@fontsource/noto-sans-devanagari/files/noto-sans-devanagari-devanagari-400-normal.woff",
  hebrew:
    "../../../../node_modules/@fontsource/noto-sans-hebrew/files/noto-sans-hebrew-hebrew-400-normal.woff",
} as const;

export type ShapingTestScript = keyof typeof SHAPING_TEST_FACES;

export const shapingTestFontPath = (script: ShapingTestScript): string =>
  join(import.meta.dir, SHAPING_TEST_FACES[script]);

/** Whether every shaping fixture is on disk, checked before any test body runs. */
export const SHAPING_TEST_FONTS_INSTALLED = Object.keys(SHAPING_TEST_FACES).every((script) =>
  existsSync(shapingTestFontPath(script as ShapingTestScript)),
);

export const SHAPING_TEST_FONTS_SKIP_REASON =
  "needs the @fontsource shaping fixtures installed (bun install)";

export const readShapingTestFont = async (script: ShapingTestScript): Promise<Uint8Array> =>
  new Uint8Array(await Bun.file(shapingTestFontPath(script)).arrayBuffer());
