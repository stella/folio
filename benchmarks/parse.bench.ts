/**
 * Parse benchmark: DOCX bytes → document model.
 *
 * Compares folio against the JS libraries that read a real `.docx`:
 *  - `@eigenpal/docx-editor-core` — the upstream folio forked, so the closest
 *    apples-to-apples (same parse → model pipeline, same `parseDocx` API).
 *  - `mammoth` — the de-facto DOCX → HTML converter. Different output (HTML, not
 *    an editable model), but the most common "read a docx" path in JS.
 *
 * Each iteration gets a fresh byte copy so a parser that retains/transfers its
 * input cannot skew the next run, and the copy cost is identical across libs.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import { parseDocx as eigenpalParse } from "@eigenpal/docx-editor-core";
import { parseDocx as folioParse } from "@stll/folio-core/docx/parser";
import mammoth from "mammoth";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";
import { FIXTURES, fixtureLabel, freshArrayBuffer, freshBuffer } from "./fixtures";

export function parseBench(): Bench {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));

  for (const fixture of FIXTURES) {
    const label = fixtureLabel(fixture);

    bench.add(`folio · ${label}`, async () => {
      await folioParse(new Uint8Array(freshArrayBuffer(fixture)));
    });
    bench.add(`eigenpal · ${label}`, async () => {
      await eigenpalParse(new Uint8Array(freshArrayBuffer(fixture)));
    });
    bench.add(`mammoth · ${label}`, async () => {
      await mammoth.convertToHtml({ buffer: freshBuffer(fixture) });
    });
  }

  return bench;
}
