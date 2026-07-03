import { isCjkCodePoint } from "../../utils/scriptSegments";

/**
 * Indices in `text` where a line may end. Latin text breaks after
 * whitespace or hyphen; CJK text can break after every ideograph so a
 * partial line is filled instead of pushing the whole run to the next line.
 */
export function findWordBreaks(text: string): number[] {
  const breaks: number[] = [];

  for (let index = 0; index < text.length; ) {
    const codePoint = text.codePointAt(index);
    if (codePoint === undefined) {
      break;
    }
    const charLength = codePoint > 0xffff ? 2 : 1;
    const char = text[index];

    if (char === " " || char === "-" || char === "\t") {
      breaks.push(index + charLength);
    } else if (isCjkCodePoint(codePoint)) {
      breaks.push(index + charLength);
    }

    index += charLength;
  }

  return breaks;
}

export function isBreakChar(char: string | undefined): boolean {
  return char === " " || char === "-" || char === "\t";
}
