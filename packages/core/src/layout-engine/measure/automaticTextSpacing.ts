import { getHorizontalScaleFactor } from "../../utils/horizontalScale";
import type { Run } from "../types";
import {
  buildRunFontStyle,
  DEFAULT_FONT_FAMILY,
  DEFAULT_FONT_SIZE,
  ptToPx,
} from "./measureHelpers";
import type { AutomaticTextSpacingPolicy, FontStyle } from "./measureTypes";

const AVERAGE_WIDTH_UNITS_PER_EM = 2048;
const AUTHORED_AVERAGE_WIDTHS = new Map<
  string,
  readonly [number, number, number, number]
>([
  ["aptos", [1148, 1202, 1150, 1204]],
  ["arial", [904, 980, 904, 980]],
  ["calibri", [1067, 1098, 1067, 1099]],
  ["calibri light", [1065, 1065, 1067, 1067]],
  ["cambria", [1260, 1228, 1111, 1198]],
  ["century", [951, 951, 951, 951]],
  ["century gothic", [995, 994, 995, 994]],
  ["courier new", [1229, 1229, 1229, 1229]],
  ["times new roman", [821, 874, 823, 844]],
]);

type TextScript = "eastAsian" | "latin" | "number" | "other";
const EAST_ASIAN_SCRIPT =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\p{Script=Yi}]/u;

export const hasAutomaticTextSpacingCandidates = (text: string): boolean =>
  EAST_ASIAN_SCRIPT.test(text) && /[0-9\p{Script=Latin}]/u.test(text);

const textScript = (character: string): TextScript => {
  const codePoint = character.codePointAt(0);
  if (codePoint === undefined || (codePoint >= 0xff00 && codePoint <= 0xffef)) {
    return "other";
  }
  if (/^[0-9]$/u.test(character)) return "number";
  if (/^\p{Script=Latin}$/u.test(character)) return "latin";
  if (EAST_ASIAN_SCRIPT.test(character)) {
    return "eastAsian";
  }
  return "other";
};

const eligibleStyle = (style: FontStyle): boolean =>
  Number.isFinite(style.fontSize ?? DEFAULT_FONT_SIZE) &&
  (style.fontSize ?? DEFAULT_FONT_SIZE) > 0 &&
  style.forceComplexScript !== true &&
  (style.letterSpacing ?? 0) === 0 &&
  getHorizontalScaleFactor(style.horizontalScale) === 1 &&
  style.fontVariant === undefined &&
  style.textTransform === undefined;

const authoredReverseGapPx = (style: FontStyle): number => {
  const authoredFamily = style.automaticTextSpacing?.reverseFontFamily;
  const averageWidths = AUTHORED_AVERAGE_WIDTHS.get(
    (authoredFamily === undefined ? style.fontFamily : authoredFamily)
      ?.trim()
      .toLowerCase() ?? "",
  );
  const average = averageWidths?.at(
    (style.bold ? 1 : 0) + (style.italic ? 2 : 0),
  );
  if (average === undefined) return 0;
  return (
    (ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE) * average) /
    AVERAGE_WIDTH_UNITS_PER_EM /
    2
  );
};

type AutomaticTextGapOptions = {
  before: { character: string; style: FontStyle };
  after: { character: string; style: FontStyle };
  policy: AutomaticTextSpacingPolicy;
};

export const automaticTextGap = ({
  before,
  after,
  policy,
}: AutomaticTextGapOptions): number => {
  if (!eligibleStyle(before.style) || !eligibleStyle(after.style)) return 0;
  const left = textScript(before.character);
  const right = textScript(after.character);
  if (left !== "eastAsian" && right !== "eastAsian") return 0;
  const other = left === "eastAsian" ? right : left;
  switch (other) {
    case "latin":
      if (!policy.latin) return 0;
      break;
    case "number":
      if (!policy.numbers) return 0;
      break;
    case "eastAsian":
    case "other":
      return 0;
    default:
      other satisfies never;
      return 0;
  }
  if (left !== "eastAsian") return authoredReverseGapPx(before.style);
  return ptToPx(before.style.fontSize ?? DEFAULT_FONT_SIZE) / 4;
};

export const automaticTextSpacingGaps = (text: string, style: FontStyle) => {
  const gaps: { offset: number; advancePx: number }[] = [];
  const policy = style.automaticTextSpacing;
  if (
    policy === undefined ||
    (!policy.latin && !policy.numbers) ||
    !hasAutomaticTextSpacingCandidates(text)
  )
    return gaps;
  let previousBase = "";
  let offset = 0;
  for (const character of text) {
    if (/^\p{M}$/u.test(character)) {
      offset += character.length;
      continue;
    }
    const advancePx = automaticTextGap({
      before: { character: previousBase, style },
      after: { character, style },
      policy,
    });
    if (advancePx > 0) gaps.push({ offset, advancePx });
    previousBase = character;
    offset += character.length;
  }
  return gaps;
};

export const automaticGapCodeUnitAdvances = (
  text: string,
  gaps: readonly { offset: number; advancePx: number }[],
) => {
  const advances = new Map<number, number>();
  for (const gap of gaps) {
    const previousUnit = text.charCodeAt(gap.offset - 1);
    const previousStart =
      previousUnit >= 0xdc00 && previousUnit <= 0xdfff
        ? gap.offset - 2
        : gap.offset - 1;
    advances.set(
      previousStart,
      (advances.get(previousStart) ?? 0) + gap.advancePx,
    );
  }
  return advances;
};

const precedingBase = (text: string, end: number): string => {
  let offset = end;
  while (offset > 0) {
    const trailing = text.charCodeAt(offset - 1);
    const start =
      trailing >= 0xdc00 && trailing <= 0xdfff ? offset - 2 : offset - 1;
    const character = text.slice(Math.max(0, start), offset);
    if (!/^\p{M}$/u.test(character)) return character;
    offset = start;
  }
  return "";
};

type AutomaticTextBoundaryOptions = {
  beforeText: string;
  beforeStyle: FontStyle;
  afterText: string;
  afterStyle: FontStyle;
  beforeEnd?: number;
  afterStart?: number;
};

export const automaticTextBoundaryGap = ({
  beforeText,
  beforeStyle,
  afterText,
  afterStyle,
  beforeEnd = beforeText.length,
  afterStart = 0,
}: AutomaticTextBoundaryOptions): number => {
  const beforePolicy = beforeStyle.automaticTextSpacing;
  const afterPolicy = afterStyle.automaticTextSpacing;
  if (beforePolicy === undefined || afterPolicy === undefined) return 0;
  const character = String.fromCodePoint(
    afterText.codePointAt(afterStart) ?? 0,
  );
  return automaticTextGap({
    before: {
      character: precedingBase(beforeText, beforeEnd),
      style: beforeStyle,
    },
    after: { character, style: afterStyle },
    policy: {
      latin: beforePolicy.latin && afterPolicy.latin,
      numbers: beforePolicy.numbers && afterPolicy.numbers,
    },
  });
};

export const automaticRunGap = (
  before: Run | undefined,
  after: Run | undefined,
): number => {
  if (before?.kind !== "text" || after?.kind !== "text") return 0;
  return automaticTextBoundaryGap({
    beforeText: before.text,
    beforeStyle: buildRunFontStyle(
      before,
      DEFAULT_FONT_FAMILY,
      DEFAULT_FONT_SIZE,
    ),
    afterText: after.text,
    afterStyle: buildRunFontStyle(
      after,
      DEFAULT_FONT_FAMILY,
      DEFAULT_FONT_SIZE,
    ),
  });
};
