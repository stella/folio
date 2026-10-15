import { panic } from "better-result";

type SliceTextOptions = {
  text: string;
  units: "codeUnits" | "codePoints" | "graphemes";
  start?: number;
  end?: number;
};

let graphemeSegmenter: Intl.Segmenter | undefined;

/** Slice in the named units; code-unit bounds clip inward to complete code points. */
export const sliceText = ({ text, units, start, end }: SliceTextOptions): string => {
  switch (units) {
    case "graphemes": {
      graphemeSegmenter ??= new Intl.Segmenter(undefined, { granularity: "grapheme" });
      return Array.from(graphemeSegmenter.segment(text), ({ segment }) => segment)
        .slice(start, end)
        .join("");
    }
    case "codePoints":
      return Array.from(text).slice(start, end).join("");
    case "codeUnits": {
      const clipped = text.slice(start, end);
      const first = clipped.charCodeAt(0);
      const last = clipped.charCodeAt(clipped.length - 1);
      const from = first >= 0xdc00 && first <= 0xdfff ? 1 : 0;
      const to = last >= 0xd800 && last <= 0xdbff ? -1 : undefined;
      return clipped.slice(from, to);
    }
    default: {
      const unhandled: never = units;
      return panic("Unknown text slice units", { units: unhandled });
    }
  }
};

/** Keep at most `limit` complete code points. */
export const truncateCodePoints = (text: string, limit: number): string =>
  sliceText({ text, units: "codePoints", end: Math.max(0, Math.trunc(limit)) });

/** Keep at most `limit` complete extended grapheme clusters. */
export const truncateGraphemes = (text: string, limit: number): string =>
  sliceText({ text, units: "graphemes", end: Math.max(0, Math.trunc(limit)) });
