import type { CompareDocxOptions } from "@stll/folio-core";

type ComparePair<Comparison> = (
  base: ArrayBuffer,
  target: ArrayBuffer,
  options: CompareDocxOptions,
) => Promise<Comparison>;

type CompareCorpusPairOptions<Comparison> = {
  compare: ComparePair<Comparison>;
  base: ArrayBuffer;
  target: ArrayBuffer;
  options: CompareDocxOptions;
};

/** Run one corpus pair under both public comparison policies. */
export const compareCorpusPairInBothModes = async <Comparison>({
  compare,
  base,
  target,
  options,
}: CompareCorpusPairOptions<Comparison>): Promise<{
  strict: Comparison;
  bestEffort: Comparison;
}> => {
  const strict = await compare(base, target, options);
  const bestEffort = await compare(base, target, { ...options, mode: "bestEffort" });
  return { strict, bestEffort };
};
