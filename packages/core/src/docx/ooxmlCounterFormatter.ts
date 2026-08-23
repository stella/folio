import type { NumberFormat } from "../types/document";

const ROMAN_PAIRS = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
] as const;

const LATIN_LOWER_SEQUENCE = [..."abcdefghijklmnopqrstuvwxyz"];
const ARABIC_ALPHA_SEQUENCE = [..."أبتثجحخدذرزسشصضطظعغفقكلمنهوي"];
const ARABIC_ABJAD_SEQUENCE = [..."أبجدهوزحطيكلمنسعفصقرشتثخذضظغ"];
const ZERO_WIDTH_NON_JOINER = "\u200c";
const MAX_REPEATED_COUNTER_LENGTH = 1024;

const toRoman = (value: number): string => {
  if (value <= 0 || value > 3999) {
    return String(value);
  }

  let remaining = value;
  let result = "";
  for (const [number, numeral] of ROMAN_PAIRS) {
    while (remaining >= number) {
      result += numeral;
      remaining -= number;
    }
  }
  return result;
};

const toRepeatedSequence = (value: number, sequence: string[]): string | undefined => {
  if (!Number.isSafeInteger(value) || value <= 0) {
    return "";
  }

  const zeroBased = value - 1;
  const character = sequence.at(zeroBased % sequence.length);
  const repetitions = Math.floor(zeroBased / sequence.length) + 1;
  if (repetitions > MAX_REPEATED_COUNTER_LENGTH) {
    return undefined;
  }
  return character?.repeat(repetitions) ?? "";
};

const toOrdinal = (value: number): string => {
  const suffixes = ["th", "st", "nd", "rd"];
  const remainder = value % 100;
  return `${value}${suffixes[(remainder - 20) % 10] ?? suffixes[remainder] ?? "th"}`;
};

/** Zero-pad a counter to `width` digits (the `decimalZero` family). */
export const padDecimal = (value: number, width: number): string => {
  if (value < 0) {
    return String(value);
  }
  return String(value).padStart(width, "0");
};

/** Format one OOXML numbering counter using Microsoft Word display semantics. */
export const formatOoxmlCounter = (value: number, format: NumberFormat | undefined): string => {
  if (!Number.isFinite(value)) {
    return "";
  }

  const resolvedFormat = format ?? "decimal";
  switch (resolvedFormat) {
    case "decimal":
    case "cardinalText":
    case "ordinalText":
    case "hex":
    case "chicago":
    case "ideographDigital":
    case "japaneseCounting":
    case "aiueo":
    case "iroha":
    case "decimalFullWidth":
    case "decimalHalfWidth":
    case "japaneseLegal":
    case "japaneseDigitalTenThousand":
    case "decimalEnclosedCircle":
    case "decimalFullWidth2":
    case "aiueoFullWidth":
    case "irohaFullWidth":
    case "ganada":
    case "chosung":
    case "decimalEnclosedFullstop":
    case "decimalEnclosedCircleChinese":
    case "ideographEnclosedCircle":
    case "ideographTraditional":
    case "ideographZodiac":
    case "ideographZodiacTraditional":
    case "taiwaneseCounting":
    case "ideographLegalTraditional":
    case "taiwaneseCountingThousand":
    case "taiwaneseDigital":
    case "chineseCounting":
    case "chineseLegalSimplified":
    case "chineseCountingThousand":
    case "koreanDigital":
    case "koreanCounting":
    case "koreanLegal":
    case "koreanDigital2":
    case "vietnameseCounting":
    case "russianLower":
    case "russianUpper":
    case "hebrew1":
    case "hebrew2":
    case "hindiVowels":
    case "hindiConsonants":
    case "hindiNumbers":
    case "hindiCounting":
    case "thaiLetters":
    case "thaiNumbers":
    case "thaiCounting":
      return String(value);
    case "decimalZero":
      return padDecimal(value, 2);
    case "decimalZero3":
      return padDecimal(value, 3);
    case "decimalZero4":
      return padDecimal(value, 4);
    case "decimalZero5":
      return padDecimal(value, 5);
    case "upperRoman":
      return toRoman(value);
    case "lowerRoman":
      return toRoman(value).toLowerCase();
    case "upperLetter":
      return (toRepeatedSequence(value, LATIN_LOWER_SEQUENCE) ?? String(value)).toUpperCase();
    case "lowerLetter":
      return toRepeatedSequence(value, LATIN_LOWER_SEQUENCE) ?? String(value);
    case "ordinal":
      return toOrdinal(value);
    case "bullet":
      return "•";
    case "none":
      return "";
    case "decimalEnclosedParen":
      return `(${value})`;
    case "numberInDash":
      return `-${value}-`;
    case "arabicAlpha": {
      const counter = toRepeatedSequence(value, ARABIC_ALPHA_SEQUENCE);
      if (counter === undefined) {
        return String(value);
      }
      return counter ? `${counter}${ZERO_WIDTH_NON_JOINER}` : "";
    }
    case "arabicAbjad": {
      const counter = toRepeatedSequence(value, ARABIC_ABJAD_SEQUENCE);
      if (counter === undefined) {
        return String(value);
      }
      return counter ? `${ZERO_WIDTH_NON_JOINER}${counter}` : "";
    }
    default:
      resolvedFormat satisfies never;
      return String(value);
  }
};
