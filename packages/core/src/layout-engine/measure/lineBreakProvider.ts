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
  /** Document replacement list: characters that may not begin a line. */
  noLineBreaksBefore?: string;
  /** Document replacement list: characters that may not end a line. */
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
const SIMPLE_BREAK_TEXT =
  /^[\p{Script=Latin}\p{Script=Cyrillic}\p{Script=Greek}\p{Number}\p{Punctuation}\p{White_Space}]*$/u;
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
const NONBREAKING_SPACES = new Set(["\u00A0", "\u2007", "\u202F"]);

// ECMA-376 kinsoku defaults are language-specific and can be replaced by
// settings.xml. This conservative common set covers punctuation excluded from
// line edges across Japanese and Chinese documents when no custom list applies.
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
  policy?.kinsoku !== false &&
  (policy?.noLineBreaksBefore !== undefined
    ? policy.noLineBreaksBefore.includes(character)
    : DEFAULT_PROHIBITED_LINE_START.has(character));

const isProhibitedLineEnd = (character: string | undefined, policy?: LineBreakPolicy): boolean =>
  character !== undefined &&
  policy?.kinsoku !== false &&
  (policy?.noLineBreaksAfter !== undefined
    ? policy.noLineBreaksAfter.includes(character)
    : DEFAULT_PROHIBITED_LINE_END.has(character));

const isCjkLineBreakParticipant = (
  character: string | undefined,
  policy?: LineBreakPolicy,
): boolean =>
  character !== undefined &&
  (CJK_SCRIPT.test(character) ||
    DEFAULT_PROHIBITED_LINE_START.has(character) ||
    DEFAULT_PROHIBITED_LINE_END.has(character) ||
    policy?.noLineBreaksBefore?.includes(character) === true ||
    policy?.noLineBreaksAfter?.includes(character) === true);

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

const findSimpleGraphemeBreaks = (text: string): number[] | undefined => {
  if (!SIMPLE_BREAK_TEXT.test(text)) {
    return undefined;
  }

  const breaks: number[] = [];
  for (let index = 0; index < text.length; index += 1) {
    const codeUnit = text.charCodeAt(index);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdfff) {
      return undefined;
    }
    if (codeUnit === 0x0d && text.charCodeAt(index + 1) === 0x0a) {
      index += 1;
    }
    breaks.push(index + 1);
  }
  return breaks;
};

const findUnicodeGraphemeBreaks = (text: string, policy?: LineBreakPolicy): number[] => {
  const simpleBreaks = findSimpleGraphemeBreaks(text);
  if (simpleBreaks !== undefined) {
    return simpleBreaks;
  }

  const breaks: number[] = [];
  for (const segment of getSegmenter(graphemeSegmenters, "grapheme", policy?.locale).segment(
    text,
  )) {
    const end = segment.index + segment.segment.length;
    breaks.push(end);
  }
  return breaks;
};

const isSimpleWhitespace = (character: string, codePoint: number): boolean =>
  codePoint === 0x20 ||
  (codePoint >= 0x09 && codePoint <= 0x0d) ||
  (codePoint > 0x7f && /\s/u.test(character));

/**
 * Find the complete set of Unicode line-break opportunities for simple text
 * without invoking `Intl.Segmenter`.
 *
 * Latin, Cyrillic, and Greek letters plus ordinary numbers, punctuation, and
 * spacing are single grapheme clusters, except CRLF, which Unicode treats as
 * one cluster. Returning `undefined` for combining marks, astral characters,
 * dictionary-segmented scripts, CJK, and other complex input keeps the full
 * Unicode provider responsible for those cases.
 */
const findSimpleBreaks = (text: string, policy?: LineBreakPolicy): number[] | undefined => {
  if (!SIMPLE_BREAK_TEXT.test(text)) {
    return undefined;
  }

  const breaks: number[] = [];
  for (let index = 0; index < text.length; index += 1) {
    const codePoint = text.charCodeAt(index);
    if (codePoint >= 0xd800 && codePoint <= 0xdfff) {
      return undefined;
    }
    if (codePoint === 0x0d && text.charCodeAt(index + 1) === 0x0a) {
      continue;
    }
    // SAFETY: index is bounded by text.length, and surrogate code units have
    // already fallen back to the complete Unicode provider above.
    const character = text[index]!;
    if (
      (isSimpleWhitespace(character, codePoint) && !NONBREAKING_SPACES.has(character)) ||
      BREAK_AFTER_CHARACTER.has(character) ||
      isLegacyEthiopicBreakCharacter(character, policy)
    ) {
      pushBreak(breaks, text, index + 1, policy);
    }
  }

  return breaks;
};

const findUnicodeBreaks = (text: string, policy?: LineBreakPolicy): number[] => {
  if (text.length === 0) {
    return [];
  }

  const simpleBreaks = findSimpleBreaks(text, policy);
  if (simpleBreaks !== undefined) {
    return simpleBreaks;
  }

  const breaks: number[] = [];
  const graphemeBreaks = findUnicodeGraphemeBreaks(text, policy);
  for (const index of graphemeBreaks) {
    const previous = previousCodePoint(text, index);
    if (
      previous !== undefined &&
      ((/\s/u.test(previous) && !NONBREAKING_SPACES.has(previous)) ||
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
      if (isCjkLineBreakParticipant(previous, policy)) {
        pushBreak(breaks, text, index, policy);
      }
    }
  }

  return [...new Set(breaks)].sort((left, right) => left - right);
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
