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
 * bytes for a code point, or the harness measures its own font plumbing
 * instead of backend divergence. A face is several binaries here (see
 * {@link SUBSET_PRIORITY}), so that agreement is per code point rather than
 * per face: the measurer and the PDF writer take the first binary whose
 * `cmap` covers the code point, and {@link bundledFontFaceCss} gives the
 * browser one `@font-face` per binary carrying the `unicode-range` that
 * routes the same code point to the same file. A face the browser resolves
 * through its own font list is a face nobody measured.
 */

import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import type { HeadlessFontRequest, HeadlessFontSource } from "../src/fonts/headlessMeasure";
import { FONT_MAPPING } from "../src/utils/fontLoader";
import { parseFontFamilyList, resolveFontFamily } from "../src/utils/fontResolver";

/** `<root>/packages/core/scripts` — resolved from this file, never from the cwd. */
const REPO_ROOT = path.resolve(import.meta.dir, "..", "..", "..");

const FONTSOURCE_DIR = path.join(REPO_ROOT, "packages", "react", "node_modules", "@fontsource");

/**
 * The subsets served for one face, highest priority first.
 *
 * `@fontsource` cuts every family into *disjoint* files, so a subset is a
 * coverage fact and not a preference: `latin-ext` carries the Czech, Slovak
 * and Polish letters but no ASCII at all (its `cmap` has no entry for `A`),
 * and `latin` carries ASCII but no `ř`. Serving one binary per face would
 * therefore mean picking a script per document and painting `.notdef` for
 * everything outside it. A face is a list instead, and a code point is served
 * by the first binary whose `cmap` covers it. `latin` leads because ASCII
 * dominates every document; `latin-ext` follows because Czech, Slovak, Polish
 * and German text is this product's common case; `arabic` is what
 * `Noto Sans Arabic` is bundled for.
 *
 * The list is short because {@link bundledFontFaceCss} inlines every served
 * binary once per family alias: `cyrillic`, `greek` and `vietnamese` sit
 * beside these files and each adds megabytes to the harness page, so a subset
 * is indexed when a fixture needs it rather than in advance.
 */
const SUBSET_PRIORITY = ["latin", "latin-ext", "arabic"] as const;

type ServedSubset = (typeof SUBSET_PRIORITY)[number];

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

// ---------------------------------------------------------------------------
// Face files
// ---------------------------------------------------------------------------

/** One served binary of a face: which subset it carries and where it lives. */
type FaceBinary = {
  readonly subset: ServedSubset;
  readonly filePath: string;
};

type FaceOptions = {
  family: string;
  weight: number;
  italic: boolean;
};

/**
 * Every served subset of a face that exists on disk, in priority order.
 * Empty when no bundled family covers the name, or when the `@fontsource`
 * packages are not installed.
 */
const faceBinaries = ({ family, weight, italic }: FaceOptions): readonly FaceBinary[] => {
  const bundled = resolveBundledFamily(family);
  if (bundled === null) return [];
  const directory = BUNDLED_FAMILY_DIRECTORIES[bundled];
  const style = italic ? "italic" : "normal";
  return SUBSET_PRIORITY.flatMap((subset) => {
    const filePath = path.join(
      FONTSOURCE_DIR,
      directory,
      "files",
      `${directory}-${subset}-${String(weight)}-${style}.woff`,
    );
    return existsSync(filePath) ? [{ subset, filePath }] : [];
  });
};

const faceCacheKey = ({ family, weight, italic }: FaceOptions): string =>
  `${family.toLowerCase()}|${String(weight)}|${italic ? "i" : "n"}`;

const weightOf = (bold: boolean): number => (bold ? BOLD_WEIGHT : REGULAR_WEIGHT);

/**
 * Read the bundled faces off disk, caching bytes per face.
 *
 * A face with no file at all returns an empty list rather than a stand-in:
 * the measurer logs the substitution and the caller reports it, where a
 * silent substitution would paginate against a font nobody chose.
 */
export const createBundledFontSource = (): HeadlessFontSource => {
  const cache = new Map<string, readonly Uint8Array[]>();

  const load = ({ family, bold, italic }: HeadlessFontRequest): readonly Uint8Array[] => {
    const face = { family, weight: weightOf(bold), italic };
    const key = faceCacheKey(face);
    const cached = cache.get(key);
    if (cached !== undefined) return cached;

    const bytes = faceBinaries(face).map(({ filePath }) => new Uint8Array(readFileSync(filePath)));
    cache.set(key, bytes);
    return bytes;
  };

  return { load };
};

// ---------------------------------------------------------------------------
// Unicode ranges
// ---------------------------------------------------------------------------

/** An inclusive `[first, last]` span of code points. */
type CodePointRange = readonly [number, number];

/**
 * `@fontsource` ships the subset ranges it cut the family with, as JSON,
 * beside the files themselves. Reading that is what keeps the browser's
 * routing and this source's routing the same table rather than two tables
 * that agree until someone edits one of them.
 */
const UNICODE_RANGE_FILE = "unicode.json";

const RANGE_TOKEN_RE = /^U\+([0-9A-Fa-f]+)(?:-([0-9A-Fa-f]+))?$/u;

const parseRangeToken = (token: string): readonly CodePointRange[] => {
  const match = RANGE_TOKEN_RE.exec(token.trim());
  if (match === null) return [];
  const first = Number.parseInt(match[1] ?? "", 16);
  const last = match[2] === undefined ? first : Number.parseInt(match[2], 16);
  if (Number.isNaN(first) || Number.isNaN(last)) return [];
  const range: CodePointRange = [first, last];
  return [range];
};

const parseRangeList = (declared: string): readonly CodePointRange[] =>
  declared.split(",").flatMap(parseRangeToken);

const readDeclaredRanges = (filePath: string): ReadonlyMap<string, readonly CodePointRange[]> => {
  const ranges = new Map<string, readonly CodePointRange[]>();
  const parsed: unknown = JSON.parse(readFileSync(filePath, "utf8"));
  if (typeof parsed !== "object" || parsed === null) return ranges;
  for (const [subset, declared] of Object.entries(parsed)) {
    if (typeof declared === "string") {
      ranges.set(subset, parseRangeList(declared));
    }
  }
  return ranges;
};

const declaredRangeCache = new Map<string, ReadonlyMap<string, readonly CodePointRange[]>>();

const declaredRanges = (directory: string): ReadonlyMap<string, readonly CodePointRange[]> => {
  const cached = declaredRangeCache.get(directory);
  if (cached !== undefined) return cached;
  const filePath = path.join(FONTSOURCE_DIR, directory, UNICODE_RANGE_FILE);
  const ranges = existsSync(filePath)
    ? readDeclaredRanges(filePath)
    : new Map<string, readonly CodePointRange[]>();
  declaredRangeCache.set(directory, ranges);
  return ranges;
};

const withoutRange = (range: CodePointRange, [start, end]: CodePointRange): CodePointRange[] => {
  const [first, last] = range;
  if (end < first || start > last) return [range];
  const remainder: CodePointRange[] = [];
  if (start > first) remainder.push([first, start - 1]);
  if (end < last) remainder.push([end + 1, last]);
  return remainder;
};

/**
 * `subject` minus everything a higher-priority subset already serves.
 *
 * The declared ranges overlap (`latin` and `latin-ext` both claim the
 * combining marks U+0304, U+0308 and U+0329), and a browser resolves an
 * overlap by declaration order while this source resolves it by priority.
 * Cutting the overlap out of the lower-priority rule makes the two agree
 * whatever order the rules are emitted in.
 */
const subtractRanges = (
  subject: readonly CodePointRange[],
  taken: readonly CodePointRange[],
): readonly CodePointRange[] => {
  let remaining = subject;
  for (const range of taken) {
    remaining = remaining.flatMap((candidate) => withoutRange(candidate, range));
  }
  return remaining;
};

const hex = (codePoint: number): string => codePoint.toString(16).toUpperCase().padStart(4, "0");

const formatRangeList = (ranges: readonly CodePointRange[]): string =>
  ranges
    .map(([first, last]) => (first === last ? `U+${hex(first)}` : `U+${hex(first)}-${hex(last)}`))
    .join(",");

// ---------------------------------------------------------------------------
// `@font-face` CSS
// ---------------------------------------------------------------------------

const FACE_VARIANTS = [
  { weight: REGULAR_WEIGHT, italic: false },
  { weight: REGULAR_WEIGHT, italic: true },
  { weight: BOLD_WEIGHT, italic: false },
  { weight: BOLD_WEIGHT, italic: true },
] as const;

const base64Cache = new Map<string, string>();

/** One file's bytes, base64 once and reused across the aliases that share it. */
const base64Of = (filePath: string): string => {
  const cached = base64Cache.get(filePath);
  if (cached !== undefined) return cached;
  const encoded = readFileSync(filePath).toString("base64");
  base64Cache.set(filePath, encoded);
  return encoded;
};

type FontFaceRuleOptions = {
  readonly family: string;
  readonly binary: FaceBinary;
  readonly weight: number;
  readonly italic: boolean;
  readonly unicodeRange: readonly CodePointRange[];
};

const fontFaceRule = ({
  family,
  binary,
  weight,
  italic,
  unicodeRange,
}: FontFaceRuleOptions): string =>
  [
    "@font-face {",
    `  font-family: "${family.replaceAll('"', '\\"')}";`,
    `  src: url(data:font/woff;base64,${base64Of(binary.filePath)}) format("woff");`,
    `  font-weight: ${String(weight)};`,
    `  font-style: ${italic ? "italic" : "normal"};`,
    `  unicode-range: ${formatRangeList(unicodeRange)};`,
    "  font-display: block;",
    "}",
  ].join("\n");

/**
 * The rules for one variant of one declared family name.
 *
 * A subset whose declared range is entirely served by a higher-priority
 * subset, or that `unicode.json` does not describe, is not declared at all: a
 * rule with no `unicode-range` covers every code point and would let the
 * browser paint from a binary the PDF resolved elsewhere, which is exactly
 * the divergence this file exists to prevent.
 */
const variantRules = ({ family, weight, italic }: FaceOptions): readonly string[] => {
  const bundled = resolveBundledFamily(family);
  if (bundled === null) return [];
  const ranges = declaredRanges(BUNDLED_FAMILY_DIRECTORIES[bundled]);
  const rules: string[] = [];
  const taken: CodePointRange[] = [];
  for (const binary of faceBinaries({ family, weight, italic })) {
    const unicodeRange = subtractRanges(ranges.get(binary.subset) ?? [], taken);
    if (unicodeRange.length === 0) continue;
    rules.push(fontFaceRule({ family, binary, weight, italic, unicodeRange }));
    taken.push(...unicodeRange);
  }
  return rules;
};

export type BundledFontFaceCssOptions = {
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
 * `@font-face` rules with the `.woff` files inlined as `data:` URLs, one rule
 * per served subset of each face.
 *
 * The bytes are repeated per family because CSS has no way to point two family
 * names at one face. That costs a few megabytes of local HTML and buys the
 * only property that matters here: the browser paints, per code point, the
 * binary the headless measurer measured, so a raster difference is a backend
 * difference.
 */
export const bundledFontFaceCss = (options: BundledFontFaceCssOptions = {}): string => {
  const declared = new Map<string, string>();
  for (const family of [...bundledFamilyAliases(), ...(options.families ?? [])]) {
    declared.set(family.trim().toLowerCase(), family.trim());
  }
  return [...declared.values()]
    .flatMap((family) =>
      FACE_VARIANTS.flatMap(({ weight, italic }) => variantRules({ family, weight, italic })),
    )
    .join("\n");
};
