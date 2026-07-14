/**
 * Replaceable line-break policy for the layout engine.
 *
 * Providers return UTF-16 offsets so their output composes directly with the
 * run and ProseMirror positions used by the rest of the layout pipeline. The
 * default provider uses the platform's Unicode segmenter, then applies the
 * OOXML line-edge restrictions that are available in the flow model.
 */

import czechHyphenation from "hyphen/cs";
import britishEnglishHyphenation from "hyphen/en-gb";
import americanEnglishHyphenation from "hyphen/en-us";
import slovakHyphenation from "hyphen/sk";

const { hyphenateSync: hyphenateCzech } = czechHyphenation;
const { hyphenateSync: hyphenateBritishEnglish } = britishEnglishHyphenation;
const { hyphenateSync: hyphenateAmericanEnglish } = americanEnglishHyphenation;
const { hyphenateSync: hyphenateSlovak } = slovakHyphenation;

export type LineBreakPolicy = {
  /** BCP-47 language tag resolved from `w:lang` for this run. */
  locale?: string;
  /** Apply East Asian first/last-character restrictions (`w:kinsoku`). */
  kinsoku?: boolean;
  /** Document override: characters that may not begin a line. */
  noLineBreaksBefore?: string;
  /** Document override: characters that may not end a line. */
  noLineBreaksAfter?: string;
  /** Use the legacy Ethiopic/Amharic compatibility behavior. */
  useLegacyEthiopicAmharicRules?: boolean;
  /** Keep words made entirely of capital letters unhyphenated. */
  doNotHyphenateCaps?: boolean;
  /** The run renders with the DOCX all-caps transform. */
  renderedAllCaps?: boolean;
};

export type LineBreakProvider = {
  /** Legal soft-wrap offsets, in ascending UTF-16 order. */
  findBreaks: (text: string, policy?: LineBreakPolicy) => number[];
  /** Grapheme-safe emergency-wrap offsets, in ascending UTF-16 order. */
  findGraphemeBreaks: (text: string, policy?: LineBreakPolicy) => number[];
  /** Dictionary-based discretionary hyphen offsets, in ascending UTF-16 order. */
  findHyphenationBreaks?: (text: string, policy?: LineBreakPolicy) => number[];
  /** Whether a trailing grapheme may overhang the line edge. */
  isHangingPunctuation?: (text: string, policy?: LineBreakPolicy) => boolean;
};

const SEGMENTER_CACHE_LIMIT = 16;
const DEFAULT_SEGMENTER_KEY = "";
const SOFT_HYPHEN = "\u00AD";
export const MAX_HYPHENATION_WORD_LENGTH = 256;

const wordSegmenters = new Map<string, Intl.Segmenter>();
const graphemeSegmenters = new Map<string, Intl.Segmenter>();

const segmenterLocale = (locale?: string): string | undefined => {
  const language = locale?.trim().replaceAll("_", "-").split("-").at(0);
  return language && /^[a-z]{2,3}$/iu.test(language) ? language.toLowerCase() : undefined;
};

const DICTIONARY_BREAK_SCRIPT =
  /[\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Myanmar}]/u;
const CJK_SCRIPT = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const BREAK_AFTER_CHARACTER = new Set([
  "-",
  "\u058A", // Armenian hyphen
  "\u05BE", // Hebrew maqaf
  "\u1400", // Canadian syllabics hyphen
  "\u1806", // Mongolian todo soft hyphen
  "\u2010", // hyphen
  "\u2E17", // double oblique hyphen
  "\u2E40", // double hyphen
  "\u30A0", // katakana-hiragana double hyphen
  "\uFE63", // small hyphen-minus
  "\uFF0D", // fullwidth hyphen-minus
  "\u200B", // zero-width space
  "\u00AD", // soft hyphen
]);
const CZECH_ONE_LETTER_PREPOSITIONS = new Set(["k", "o", "s", "u", "v", "z"]);

// ECMA-376 kinsoku defaults are language-specific and can be overridden by
// settings.xml. This conservative common set covers punctuation excluded from
// line edges across Japanese and Chinese documents; document overrides are
// layered on top below.
const DEFAULT_PROHIBITED_LINE_START = new Set([
  "!",
  "%",
  ")",
  ",",
  ".",
  ":",
  ";",
  "?",
  "]",
  "}",
  "、",
  "。",
  "〉",
  "》",
  "」",
  "』",
  "】",
  "〕",
  "）",
  "］",
  "｝",
  "，",
  "．",
  "：",
  "；",
  "！",
  "？",
  "…",
  "’",
  "”",
]);

const DEFAULT_PROHIBITED_LINE_END = new Set([
  "$",
  "(",
  "[",
  "{",
  "£",
  "¥",
  "〈",
  "《",
  "「",
  "『",
  "【",
  "〔",
  "（",
  "［",
  "｛",
  "‘",
  "“",
]);

const getSegmenter = (
  cache: Map<string, Intl.Segmenter>,
  granularity: "word" | "grapheme",
  locale?: string,
): Intl.Segmenter => {
  const normalizedLocale = segmenterLocale(locale);
  const key = normalizedLocale ?? DEFAULT_SEGMENTER_KEY;
  const cached = cache.get(key);
  if (cached) {
    return cached;
  }

  const segmenter = new Intl.Segmenter(normalizedLocale, { granularity });
  if (cache.size >= SEGMENTER_CACHE_LIMIT) {
    const oldestKey = cache.keys().next().value;
    if (typeof oldestKey === "string") {
      cache.delete(oldestKey);
    }
  }
  cache.set(key, segmenter);
  return segmenter;
};

const firstCodePoint = (text: string, index: number): string | undefined => {
  if (index >= text.length) {
    return undefined;
  }
  return String.fromCodePoint(text.codePointAt(index) ?? 0);
};

const previousCodePoint = (text: string, index: number): string | undefined => {
  if (index <= 0) {
    return undefined;
  }
  const trailing = text.charCodeAt(index - 1);
  if (trailing >= 0xdc00 && trailing <= 0xdfff && index >= 2) {
    return text.slice(index - 2, index);
  }
  return text[index - 1];
};

const isProhibitedLineStart = (character: string | undefined, policy?: LineBreakPolicy): boolean =>
  character !== undefined &&
  (policy?.noLineBreaksBefore?.includes(character) === true ||
    (policy?.kinsoku !== false && DEFAULT_PROHIBITED_LINE_START.has(character)));

const isProhibitedLineEnd = (character: string | undefined, policy?: LineBreakPolicy): boolean =>
  character !== undefined &&
  (policy?.noLineBreaksAfter?.includes(character) === true ||
    (policy?.kinsoku !== false && DEFAULT_PROHIBITED_LINE_END.has(character)));

const isLegacyEthiopicBreakCharacter = (
  character: string | undefined,
  policy?: LineBreakPolicy,
): boolean => {
  if (!policy?.useLegacyEthiopicAmharicRules || character === undefined) {
    return false;
  }
  const codePoint = character.codePointAt(0);
  return codePoint !== undefined && codePoint >= 0x1361 && codePoint <= 0x1368;
};

const czechProtectedBreaks = (text: string, policy?: LineBreakPolicy): Set<number> => {
  const breaks = new Set<number>();
  if (segmenterLocale(policy?.locale) !== "cs") {
    return breaks;
  }

  let tokenStart = 0;
  let index = 0;
  while (index < text.length) {
    const character = firstCodePoint(text, index);
    if (character === undefined) {
      break;
    }
    if (!/\s/u.test(character)) {
      index += character.length;
      continue;
    }

    const token = text.slice(tokenStart, index);
    const protectsNextWord =
      token.length === 1 && CZECH_ONE_LETTER_PREPOSITIONS.has(token.toLocaleLowerCase("cs"));
    while (index < text.length) {
      const whitespace = firstCodePoint(text, index);
      if (whitespace === undefined || !/\s/u.test(whitespace)) {
        break;
      }
      index += whitespace.length;
      if (protectsNextWord) {
        breaks.add(index);
      }
    }
    tokenStart = index;
  }
  return breaks;
};

const allowsBreak = (text: string, index: number, policy?: LineBreakPolicy): boolean => {
  const previous = previousCodePoint(text, index);
  const next = firstCodePoint(text, index);
  return !isProhibitedLineEnd(previous, policy) && !isProhibitedLineStart(next, policy);
};

const pushBreak = (
  breaks: number[],
  text: string,
  index: number,
  policy?: LineBreakPolicy,
): void => {
  if (index <= 0 || index > text.length || breaks.at(-1) === index) {
    return;
  }
  if (allowsBreak(text, index, policy)) {
    breaks.push(index);
  }
};

const findUnicodeGraphemeBreaks = (text: string, policy?: LineBreakPolicy): number[] => {
  const breaks: number[] = [];
  for (const segment of getSegmenter(graphemeSegmenters, "grapheme", policy?.locale).segment(
    text,
  )) {
    const end = segment.index + segment.segment.length;
    breaks.push(end);
  }
  return breaks;
};

const findUnicodeBreaks = (text: string, policy?: LineBreakPolicy): number[] => {
  if (text.length === 0) {
    return [];
  }

  const breaks: number[] = [];
  const graphemeBreaks = findUnicodeGraphemeBreaks(text, policy);
  for (const index of graphemeBreaks) {
    const previous = previousCodePoint(text, index);
    if (
      previous !== undefined &&
      (/\s/u.test(previous) ||
        BREAK_AFTER_CHARACTER.has(previous) ||
        isLegacyEthiopicBreakCharacter(previous, policy))
    ) {
      pushBreak(breaks, text, index, policy);
    }
  }

  const wordSegments = [...getSegmenter(wordSegmenters, "word", policy?.locale).segment(text)];

  for (const [segmentIndex, segment] of wordSegments.entries()) {
    const end = segment.index + segment.segment.length;
    const nextSegment = wordSegments[segmentIndex + 1];

    if (
      segment.isWordLike &&
      nextSegment?.isWordLike === true &&
      DICTIONARY_BREAK_SCRIPT.test(segment.segment + nextSegment.segment)
    ) {
      pushBreak(breaks, text, end, policy);
    }
  }

  if (CJK_SCRIPT.test(text)) {
    for (const index of graphemeBreaks) {
      const previous = previousCodePoint(text, index);
      const next = firstCodePoint(text, index);
      if (
        (previous !== undefined && CJK_SCRIPT.test(previous)) ||
        (isProhibitedLineStart(previous, policy) && next !== undefined && CJK_SCRIPT.test(next))
      ) {
        pushBreak(breaks, text, index, policy);
      }
    }
  }

  const protectedBreaks = czechProtectedBreaks(text, policy);
  return [...new Set(breaks)]
    .filter((index) => !protectedBreaks.has(index))
    .sort((left, right) => left - right);
};

type HyphenateWord = (text: string) => string;

const hyphenatorFor = (locale?: string): HyphenateWord | undefined => {
  const normalized = locale?.trim().replaceAll("_", "-").toLowerCase();
  if (!normalized) {
    return undefined;
  }
  if (normalized === "en-gb" || normalized.startsWith("en-gb-")) {
    return hyphenateBritishEnglish;
  }
  if (normalized === "en-us" || normalized.startsWith("en-us-")) {
    return hyphenateAmericanEnglish;
  }
  if (normalized === "cs" || normalized.startsWith("cs-")) {
    return hyphenateCzech;
  }
  if (normalized === "sk" || normalized.startsWith("sk-")) {
    return hyphenateSlovak;
  }
  return undefined;
};

const isAllCapsWord = (text: string, locale?: string): boolean => {
  const letters = [...text].filter((character) => /\p{Letter}/u.test(character)).join("");
  if (letters.length === 0) {
    return false;
  }
  const casingLocale = segmenterLocale(locale);
  return (
    letters === letters.toLocaleUpperCase(casingLocale) &&
    letters !== letters.toLocaleLowerCase(casingLocale)
  );
};

const findPatternHyphenationBreaks = (text: string, policy?: LineBreakPolicy): number[] => {
  if (text.length > MAX_HYPHENATION_WORD_LENGTH || text.includes(SOFT_HYPHEN)) {
    return [];
  }
  const hyphenate = hyphenatorFor(policy?.locale);
  if (!hyphenate) {
    return [];
  }
  if (
    policy?.doNotHyphenateCaps &&
    (policy.renderedAllCaps === true || isAllCapsWord(text, policy.locale))
  ) {
    return [];
  }

  const hyphenated = hyphenate(text);
  const breaks: number[] = [];
  let sourceOffset = 0;
  for (const character of hyphenated) {
    if (character === SOFT_HYPHEN) {
      breaks.push(sourceOffset);
      continue;
    }
    sourceOffset += character.length;
  }
  return sourceOffset === text.length ? breaks : [];
};

const isDefaultHangingPunctuation = (text: string, policy?: LineBreakPolicy): boolean => {
  const characters = [...text];
  return (
    characters.length > 0 &&
    characters.every((character) => isProhibitedLineStart(character, policy))
  );
};

export const defaultLineBreakProvider: LineBreakProvider = {
  findBreaks: findUnicodeBreaks,
  findGraphemeBreaks: findUnicodeGraphemeBreaks,
  findHyphenationBreaks: findPatternHyphenationBreaks,
  isHangingPunctuation: isDefaultHangingPunctuation,
};

let activeLineBreakProvider = defaultLineBreakProvider;
let lineBreakProviderGeneration = 0;

export const getLineBreakProvider = (): LineBreakProvider => activeLineBreakProvider;

/** Changes whenever the active provider changes, so layout caches cannot go stale. */
export const getLineBreakProviderGeneration = (): number => lineBreakProviderGeneration;

export const setLineBreakProvider = (provider: LineBreakProvider): void => {
  activeLineBreakProvider = provider;
  lineBreakProviderGeneration += 1;
};

export const resetLineBreakProvider = (): void => {
  activeLineBreakProvider = defaultLineBreakProvider;
  lineBreakProviderGeneration += 1;
};
