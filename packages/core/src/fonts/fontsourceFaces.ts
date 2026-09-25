/**
 * Faces from the `@fontsource` packages, for a caller that measures and paints
 * without a browser's font stack: a PDF export, a server-rendered preview, a
 * paint harness.
 *
 * One table answers both consumers. The headless measurer and the PDF writer
 * load binaries through {@link FontsourceFaces.source}; a page painted by a
 * browser declares the same binaries through {@link FontsourceFaces.fontFaceCss},
 * each subset under the `unicode-range` `@fontsource` cut it with. Routing a
 * code point to one binary on both sides is what keeps the painted page the
 * page that was measured.
 *
 * The module reads no files itself: the caller supplies
 * {@link FontsourceFiles.read}, so it runs wherever the bytes can be found (a
 * package's `node_modules`, a repository checkout, a fetch in a browser).
 */

import type { HeadlessFontRequest, HeadlessFontSource } from "./headlessMeasure";
import { bytesToDataUrl } from "../utils/base64";
import { FONT_MAPPING } from "../utils/fontLoader";
import { parseFontFamilyList, resolveFontFamily } from "../utils/fontResolver";

/** Where `@fontsource` files come from: `null` when the package or file is absent. */
export type FontsourceFiles = {
  /** Bytes of `@fontsource/<packageName>/<relativePath>`. */
  readonly read: (packageName: string, relativePath: string) => Uint8Array | null;
};

/**
 * The subsets served, in resolution priority: a code point two subsets cover
 * resolves to the earlier one on both sides.
 */
const SUBSET_PRIORITY = ["latin", "latin-ext", "arabic"] as const;

type ServedSubset = (typeof SUBSET_PRIORITY)[number];

/** Folio paints 400 and 700 only, as `DisplayFontFace` states. */
const REGULAR_WEIGHT = 400;
const BOLD_WEIGHT = 700;

/** Each bundled family and the `@fontsource` package that ships it. */
export const FONTSOURCE_PACKAGES = {
  Arimo: "arimo",
  Caladea: "caladea",
  Carlito: "carlito",
  Cousine: "cousine",
  Lato: "lato",
  "Noto Sans Arabic": "noto-sans-arabic",
  "Source Sans 3": "source-sans-3",
  Tinos: "tinos",
} as const;

type BundledFamily = keyof typeof FONTSOURCE_PACKAGES;

/** A CSS stack ends in a generic, so a generic must resolve to a bundled family too. */
const GENERIC_FAMILIES = {
  serif: "Tinos",
  "sans-serif": "Arimo",
  monospace: "Cousine",
} as const satisfies Record<string, BundledFamily>;

const isBundledFamily = (name: string): name is BundledFamily =>
  Object.hasOwn(FONTSOURCE_PACKAGES, name);

const isGenericFamily = (name: string): name is keyof typeof GENERIC_FAMILIES =>
  Object.hasOwn(GENERIC_FAMILIES, name);

/** Every name a bundled family answers to: its own, and the authored names mapped to it. */
const familyIndexEntries = (): readonly (readonly [string, BundledFamily])[] => {
  const entries: (readonly [string, BundledFamily])[] = [];
  for (const family of Object.keys(FONTSOURCE_PACKAGES)) {
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
  for (const [generic, bundled] of Object.entries(GENERIC_FAMILIES)) {
    entries.push([generic, bundled]);
  }
  return entries;
};

const FAMILY_INDEX: ReadonlyMap<string, BundledFamily> = new Map(familyIndexEntries());

/** Every family name the faces answer to, in a stable order for the CSS. */
const bundledFamilyAliases = (): readonly string[] =>
  [
    ...Object.keys(FONTSOURCE_PACKAGES),
    ...Object.keys(FONT_MAPPING).filter((authored) => FAMILY_INDEX.has(authored.toLowerCase())),
  ].sort((a, b) => a.localeCompare(b, "en"));

/**
 * The bundled family a requested name resolves to: directly, or through the
 * same fallback chain the editor's font resolver walks.
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

type FaceOptions = {
  family: string;
  weight: number;
  italic: boolean;
};

/** One served binary of a face: which subset it carries and its bytes. */
type FaceBinary = {
  readonly subset: ServedSubset;
  readonly bytes: Uint8Array;
};

const weightOf = (bold: boolean): number => (bold ? BOLD_WEIGHT : REGULAR_WEIGHT);

/** An inclusive `[first, last]` span of code points. */
type CodePointRange = readonly [number, number];

/** `@fontsource` ships the ranges it cut each subset with, beside the files. */
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

const parseDeclaredRanges = (bytes: Uint8Array): ReadonlyMap<string, readonly CodePointRange[]> => {
  const ranges = new Map<string, readonly CodePointRange[]>();
  const parsed: unknown = JSON.parse(new TextDecoder().decode(bytes));
  if (typeof parsed !== "object" || parsed === null) return ranges;
  for (const [subset, declared] of Object.entries(parsed)) {
    if (typeof declared === "string") {
      ranges.set(subset, parseRangeList(declared));
    }
  }
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
 * `subject` minus everything a higher-priority subset already serves. The
 * declared ranges overlap (`latin` and `latin-ext` both claim some combining
 * marks); cutting the overlap out of the lower-priority rule makes a browser,
 * which resolves an overlap by declaration order, agree with this table.
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

/**
 * A CSS `<string>` token, quoted and escaped. A family name comes from an
 * authored document and is emitted into a `<style>` element, so the
 * backslash is escaped first, then the quote, and line terminators become hex
 * escapes, which CSS strings require.
 */
export const cssString = (value: string): string =>
  `"${value
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\n", "\\A ")
    .replaceAll("\r", "\\D ")
    .replaceAll("\f", "\\C ")}"`;

const FACE_VARIANTS = [
  { weight: REGULAR_WEIGHT, italic: false },
  { weight: REGULAR_WEIGHT, italic: true },
  { weight: BOLD_WEIGHT, italic: false },
  { weight: BOLD_WEIGHT, italic: true },
] as const;

/** The media type a `.woff` face is inlined under. */
const WOFF_MEDIA_TYPE = "font/woff";

type FontFaceRuleOptions = {
  readonly family: string;
  readonly dataUrl: string;
  readonly weight: number;
  readonly italic: boolean;
  readonly unicodeRange: readonly CodePointRange[];
};

const fontFaceRule = ({ family, dataUrl, weight, italic, unicodeRange }: FontFaceRuleOptions) =>
  [
    "@font-face {",
    `  font-family: ${cssString(family)};`,
    `  src: url(${dataUrl}) format("woff");`,
    `  font-weight: ${String(weight)};`,
    `  font-style: ${italic ? "italic" : "normal"};`,
    `  unicode-range: ${formatRangeList(unicodeRange)};`,
    "  font-display: block;",
    "}",
  ].join("\n");

export type FontsourceFontFaceCssOptions = {
  /**
   * Families to declare beyond the names the faces answer to directly. Pass a
   * display list's own `fonts` families: a face reached through the fallback
   * chain has no rule of its own, and the page names it by its authored name.
   */
  readonly families?: readonly string[];
};

export type FontsourceFaces = {
  /** Binaries for the headless measurer and the PDF writer. */
  readonly source: HeadlessFontSource;
  /**
   * `@font-face` rules with the `.woff` bytes inlined as `data:` URLs, one
   * rule per served subset of each face, so a browser paints each code point
   * from the binary that was measured.
   */
  readonly fontFaceCss: (options?: FontsourceFontFaceCssOptions) => string;
};

/** The `@fontsource` faces, read through `files` and cached per face. */
export const createFontsourceFaces = (files: FontsourceFiles): FontsourceFaces => {
  const faceCache = new Map<string, readonly FaceBinary[]>();
  const rangeCache = new Map<string, ReadonlyMap<string, readonly CodePointRange[]>>();
  const dataUrlCache = new Map<Uint8Array, string>();

  const faceBinaries = ({ family, weight, italic }: FaceOptions): readonly FaceBinary[] => {
    const bundled = resolveBundledFamily(family);
    if (bundled === null) return [];
    const packageName = FONTSOURCE_PACKAGES[bundled];
    const style = italic ? "italic" : "normal";
    const key = `${packageName}|${String(weight)}|${style}`;
    const cached = faceCache.get(key);
    if (cached !== undefined) return cached;
    const binaries = SUBSET_PRIORITY.flatMap((subset) => {
      const bytes = files.read(
        packageName,
        `files/${packageName}-${subset}-${String(weight)}-${style}.woff`,
      );
      return bytes === null ? [] : [{ subset, bytes }];
    });
    faceCache.set(key, binaries);
    return binaries;
  };

  const declaredRanges = (packageName: string) => {
    const cached = rangeCache.get(packageName);
    if (cached !== undefined) return cached;
    const bytes = files.read(packageName, UNICODE_RANGE_FILE);
    const ranges =
      bytes === null ? new Map<string, readonly CodePointRange[]>() : parseDeclaredRanges(bytes);
    rangeCache.set(packageName, ranges);
    return ranges;
  };

  const dataUrlOf = (bytes: Uint8Array): string => {
    const cached = dataUrlCache.get(bytes);
    if (cached !== undefined) return cached;
    const encoded = bytesToDataUrl(bytes, WOFF_MEDIA_TYPE);
    dataUrlCache.set(bytes, encoded);
    return encoded;
  };

  /**
   * The rules for one variant of one family name. A subset whose range a
   * higher-priority subset already serves, or that `unicode.json` does not
   * describe, is not declared: a rule with no `unicode-range` would cover
   * every code point.
   */
  const variantRules = ({ family, weight, italic }: FaceOptions): readonly string[] => {
    const bundled = resolveBundledFamily(family);
    if (bundled === null) return [];
    const ranges = declaredRanges(FONTSOURCE_PACKAGES[bundled]);
    const rules: string[] = [];
    const taken: CodePointRange[] = [];
    for (const binary of faceBinaries({ family, weight, italic })) {
      const unicodeRange = subtractRanges(ranges.get(binary.subset) ?? [], taken);
      if (unicodeRange.length === 0) continue;
      rules.push(
        fontFaceRule({ family, dataUrl: dataUrlOf(binary.bytes), weight, italic, unicodeRange }),
      );
      taken.push(...unicodeRange);
    }
    return rules;
  };

  return {
    // A face with no file returns an empty list rather than a stand-in: the
    // measurer records the substitution and the caller reports it.
    source: {
      load: ({ family, bold, italic }: HeadlessFontRequest) =>
        faceBinaries({ family, weight: weightOf(bold), italic }).map(({ bytes }) => bytes),
    },
    fontFaceCss: (options = {}) => {
      const declared = new Map<string, string>();
      for (const family of [...bundledFamilyAliases(), ...(options.families ?? [])]) {
        declared.set(family.trim().toLowerCase(), family.trim());
      }
      return [...declared.values()]
        .flatMap((family) =>
          FACE_VARIANTS.flatMap(({ weight, italic }) => variantRules({ family, weight, italic })),
        )
        .join("\n");
    },
  };
};
