/**
 * A {@link HeadlessFontSource} over the `@fontsource` families folio bundles.
 *
 * Those packages are a dependency of `packages/react`, not of
 * `packages/core`, and must not become one: the core package ships no font
 * binaries. This module is tooling — `packages/core/tsconfig.build.json`
 * excludes `scripts`, so nothing here reaches the published build — which is
 * why resolving the faces off disk is fine here and an npm dependency on
 * `packages/core` would not be.
 *
 * ## Why WOFF and not WOFF2
 *
 * `fonts/sfnt/woff.ts` decodes WOFF 1.0, a per-table zlib repackaging. WOFF2
 * reverses a `glyf`/`loca` transform on top of Brotli and is rejected there,
 * so the `.woff` file is the one this source can actually parse.
 *
 * ## One resolution table, three consumers
 *
 * The measurer, the PDF writer and the browser page must all end at the same
 * bytes for a family, or the harness measures its own font plumbing instead of
 * backend divergence. {@link bundledFontFaceCss} therefore emits an
 * `@font-face` for every name {@link resolveBundledFamily} accepts — the Word
 * names included — pointing at the very file the measurer read. A face the
 * browser resolves through its own font list is a face nobody measured.
 */

import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import type { HeadlessFontRequest, HeadlessFontSource } from "../src/fonts/headlessMeasure";
import { FONT_MAPPING } from "../src/utils/fontLoader";
import { parseFontFamilyList, resolveFontFamily } from "../src/utils/fontResolver";

/** `<root>/packages/core/scripts` — resolved from this file, never from the cwd. */
const REPO_ROOT = path.resolve(import.meta.dir, "..", "..", "..");

const FONTSOURCE_DIR = path.join(REPO_ROOT, "packages", "react", "node_modules", "@fontsource");

/** The subset folio's fixtures are written in. */
const DEFAULT_SUBSET = "latin";

/** Folio paints 400 and 700 only, as `DisplayFontFace` states. */
const REGULAR_WEIGHT = 400;
const BOLD_WEIGHT = 700;

/**
 * Bundled family name to its `@fontsource` package directory. Keyed by the
 * family a document may name directly; the Word names reach these through
 * {@link FONT_MAPPING}.
 */
const BUNDLED_FAMILY_DIRECTORIES = {
  Arimo: "arimo",
  Caladea: "caladea",
  Carlito: "carlito",
  Cousine: "cousine",
  Lato: "lato",
  "Noto Sans Arabic": "noto-sans-arabic",
  "Source Sans 3": "source-sans-3",
  Tinos: "tinos",
} as const;

type BundledFamily = keyof typeof BUNDLED_FAMILY_DIRECTORIES;

/**
 * Generic categories {@link FONT_MAPPING} falls back to, pinned to a bundled
 * family. A generic left to the browser would paint a face the measurer never
 * saw, which is the one divergence this harness must not introduce itself.
 */
const GENERIC_FAMILIES = {
  serif: "Tinos",
  "sans-serif": "Arimo",
  monospace: "Cousine",
} as const satisfies Record<string, BundledFamily>;

const isBundledFamily = (name: string): name is BundledFamily =>
  Object.hasOwn(BUNDLED_FAMILY_DIRECTORIES, name);

const isGenericFamily = (name: string): name is keyof typeof GENERIC_FAMILIES =>
  Object.hasOwn(GENERIC_FAMILIES, name);

/**
 * Case-insensitive index over every name this source answers to: the bundled
 * families themselves, plus the Word families `FONT_MAPPING` substitutes.
 * Mirroring that map rather than restating it is what keeps the harness
 * painting the faces the editor paints.
 */
const familyIndexEntries = (): readonly (readonly [string, BundledFamily])[] => {
  const entries: (readonly [string, BundledFamily])[] = [];
  for (const family of Object.keys(BUNDLED_FAMILY_DIRECTORIES)) {
    if (isBundledFamily(family)) {
      entries.push([family.toLowerCase(), family]);
    }
  }
  for (const [authored, mapped] of Object.entries(FONT_MAPPING)) {
    if (isBundledFamily(mapped)) {
      entries.push([authored.toLowerCase(), mapped]);
      continue;
    }
    if (isGenericFamily(mapped)) {
      entries.push([authored.toLowerCase(), GENERIC_FAMILIES[mapped]]);
    }
  }
  // A CSS stack ends in a generic, so the walk below must be able to answer
  // for one or it would stop one entry short of where a browser stops.
  for (const [generic, bundled] of Object.entries(GENERIC_FAMILIES)) {
    entries.push([generic, bundled]);
  }
  return entries;
};

const FAMILY_INDEX: ReadonlyMap<string, BundledFamily> = new Map(familyIndexEntries());

/** Every family name the source answers to, in a stable order for the CSS. */
const bundledFamilyAliases = (): readonly string[] =>
  [
    ...Object.keys(BUNDLED_FAMILY_DIRECTORIES),
    ...Object.keys(FONT_MAPPING).filter((authored) => FAMILY_INDEX.has(authored.toLowerCase())),
  ].sort((a, b) => a.localeCompare(b, "en"));

/**
 * The bundled family a name resolves to, or null when nothing covers it.
 *
 * A name with no bundled face of its own walks folio's *own* fallback chain,
 * the one `resolveFontFamily` emits and a browser follows character by
 * character. Answering for the authored family alone was a divergence rather
 * than a gap: the editor paints Corbel with Arimo (Arial's bundled
 * substitute, second in Corbel's stack) while a null here sends the PDF
 * writer to a base-14 stand-in, so one display list reached two different
 * faces. Walking the same list is what makes the two arms agree.
 */
const resolveBundledFamily = (family: string): BundledFamily | null => {
  const direct = FAMILY_INDEX.get(family.trim().toLowerCase());
  if (direct !== undefined) {
    return direct;
  }
  for (const candidate of parseFontFamilyList(resolveFontFamily(family).cssFallback)) {
    const resolved = FAMILY_INDEX.get(candidate.trim().toLowerCase());
    if (resolved !== undefined) {
      return resolved;
    }
  }
  return null;
};

type FaceFileOptions = {
  family: string;
  weight: number;
  italic: boolean;
  subset: string;
};

const faceFilePath = ({ family, weight, italic, subset }: FaceFileOptions): string | null => {
  const bundled = resolveBundledFamily(family);
  if (bundled === null) return null;
  const directory = BUNDLED_FAMILY_DIRECTORIES[bundled];
  const style = italic ? "italic" : "normal";
  const filePath = path.join(
    FONTSOURCE_DIR,
    directory,
    "files",
    `${directory}-${subset}-${String(weight)}-${style}.woff`,
  );
  return existsSync(filePath) ? filePath : null;
};

const faceCacheKey = ({ family, weight, italic }: Omit<FaceFileOptions, "subset">): string =>
  `${family.toLowerCase()}|${String(weight)}|${italic ? "i" : "n"}`;

const weightOf = (bold: boolean): number => (bold ? BOLD_WEIGHT : REGULAR_WEIGHT);

export type BundledFontSourceOptions = {
  /**
   * `latin`, `latin-ext`, `cyrillic`, `greek`, `hebrew`, `vietnamese`, ...
   *
   * One subset, not a union: `load` returns one binary per face and the
   * `@fontsource` split is disjoint, so a subset is a real limit rather than a
   * preference. `latin-ext` carries the Czech and Polish letters but *no
   * ASCII* (its `cmap` has no entry for `A`), so preferring it would leave
   * every ordinary word painting `.notdef`. `latin` is therefore the default,
   * and a document outside it is a reported substitution, not a silent one.
   * {@link bundledFontFaceCss} declares exactly this subset with no
   * `unicode-range`, so the browser cannot reach a face the PDF did not embed.
   */
  readonly subset?: string;
};

/**
 * Read the bundled faces off disk, caching bytes per face.
 *
 * A face with no file returns `null` rather than a stand-in: the measurer logs
 * the substitution and the caller reports it, where a silent substitution
 * would paginate against a font nobody chose.
 */
export const createBundledFontSource = (
  options: BundledFontSourceOptions = {},
): HeadlessFontSource => {
  const subset = options.subset ?? DEFAULT_SUBSET;
  const cache = new Map<string, Uint8Array | null>();

  const load = ({ family, bold, italic }: HeadlessFontRequest): Uint8Array | null => {
    const weight = weightOf(bold);
    const key = faceCacheKey({ family, weight, italic });
    const cached = cache.get(key);
    if (cached !== undefined) return cached;

    const filePath = faceFilePath({ family, weight, italic, subset });
    const bytes = filePath === null ? null : new Uint8Array(readFileSync(filePath));
    cache.set(key, bytes);
    return bytes;
  };

  return { load };
};

const FACE_VARIANTS = [
  { weight: REGULAR_WEIGHT, italic: false },
  { weight: REGULAR_WEIGHT, italic: true },
  { weight: BOLD_WEIGHT, italic: false },
  { weight: BOLD_WEIGHT, italic: true },
] as const;

const fontFaceRule = (family: string, options: FaceFileOptions): string | null => {
  const filePath = faceFilePath(options);
  if (filePath === null) return null;
  const base64 = readFileSync(filePath).toString("base64");
  return [
    "@font-face {",
    `  font-family: "${family.replaceAll('"', '\\"')}";`,
    `  src: url(data:font/woff;base64,${base64}) format("woff");`,
    `  font-weight: ${String(options.weight)};`,
    `  font-style: ${options.italic ? "italic" : "normal"};`,
    "  font-display: block;",
    "}",
  ].join("\n");
};

export type BundledFontFaceCssOptions = BundledFontSourceOptions & {
  /**
   * Families to declare on top of the ones this source answers to by name.
   * Pass a display list's own `fonts` families: a face the source resolves
   * through the fallback chain (Corbel to Arimo) has no `@font-face` of its
   * own, and the DOM backend emits `"Corbel", serif`, so without a rule the
   * browser would paint its own serif while the PDF embedded Arimo.
   */
  readonly families?: readonly string[];
};

/**
 * `@font-face` rules with the `.woff` files inlined as `data:` URLs.
 *
 * The bytes are repeated per family because CSS has no way to point two family
 * names at one face. That costs a few megabytes of local HTML and buys the
 * only property that matters here: the browser paints the faces the headless
 * measurer measured, so a raster difference is a backend difference.
 */
export const bundledFontFaceCss = (options: BundledFontFaceCssOptions = {}): string => {
  const subset = options.subset ?? DEFAULT_SUBSET;
  const declared = new Map<string, string>();
  for (const family of [...bundledFamilyAliases(), ...(options.families ?? [])]) {
    declared.set(family.trim().toLowerCase(), family.trim());
  }
  return [...declared.values()]
    .flatMap((family) =>
      FACE_VARIANTS.flatMap(({ weight, italic }) => {
        const rule = fontFaceRule(family, { family, weight, italic, subset });
        return rule === null ? [] : [rule];
      }),
    )
    .join("\n");
};
