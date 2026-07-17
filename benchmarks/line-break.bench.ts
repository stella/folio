/**
 * Line-break benchmark: representative paragraph text → legal wrap offsets.
 *
 * The simple corpus tracks the common fast path without reducing the default
 * provider's Unicode coverage. The mixed corpus keeps the segmenter-backed
 * CJK, right-to-left, and international fallback visible alongside it.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import {
  findGraphemeBreaks,
  findWordBreaks,
} from "@stll/folio-core/layout-engine/measure/lineBreaks";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";

const SIMPLE_PARAGRAPHS = Array.from(
  { length: 1_500 },
  (_, index) =>
    `Performance paragraph ${index + 1}: repository-authored text exercises DOCX parsing, conversion, measurement, pagination, and painting.`,
);
const MIXED_PARAGRAPHS = Array.from({ length: 250 }, (_, index) => [
  `${index + 1}. English and Latin: Performance baseline for international documents.`,
  `${index + 1}. العربية: هذا مستند اصطناعي لقياس أداء التخطيط والتحرير.`,
  `${index + 1}. עברית: זהו מסמך סינתטי למדידת ביצועי פריסה ועריכה.`,
  `${index + 1}. 中文與日本語: 這是一個用於效能測量的合成文件。日本語の文章も含みます。`,
  `${index + 1}. Mixed: Contract סעיף 12 يتضمن شروطًا متعددة，並包含 multilingual text.`,
]).flat();

export function lineBreakBench(): Bench {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));

  bench.add("simple Latin · 1500 paragraphs", () => countBreaks(SIMPLE_PARAGRAPHS));
  bench.add("simple Latin graphemes · 1500 paragraphs", () =>
    countBreaks(SIMPLE_PARAGRAPHS, findGraphemeBreaks),
  );
  bench.add("mixed script · 1250 paragraphs", () => countBreaks(MIXED_PARAGRAPHS));

  return bench;
}

const countBreaks = (
  paragraphs: readonly string[],
  findBreaks: (text: string) => number[] = findWordBreaks,
): number => {
  let count = 0;
  for (const paragraph of paragraphs) {
    count += findBreaks(paragraph).length;
  }
  return count;
};
