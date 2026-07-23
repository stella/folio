/**
 * Parse benchmark: DOCX bytes → document model.
 *
 * Each iteration gets a fresh byte copy so retained or transferred input
 * cannot skew the next run.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import { parseDocx as folioParse } from "@stll/folio-core/docx/parser";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";
import { FIXTURES, fixtureLabel, freshArrayBuffer } from "./fixtures";

export function parseBench(): Bench {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));

  for (const fixture of FIXTURES) {
    const label = fixtureLabel(fixture);

    bench.add(`folio · ${label}`, async () => {
      await folioParse(new Uint8Array(freshArrayBuffer(fixture)));
    });
  }

  return bench;
}
