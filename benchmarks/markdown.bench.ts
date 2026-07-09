/**
 * Markdown benchmark: folio's DOCX-model ↔ Markdown bridge.
 *
 * folio-only — no other JS library does the equivalent round-trip. The fixture
 * is parsed once (setup, not timed); we bench `toMarkdown` and `fromMarkdown`.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import { parseDocx as folioParse } from "@stll/folio-core/docx/parser";
import { fromMarkdown, toMarkdown } from "@stll/folio-core/markdown";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";
import { FIXTURES, fixtureLabel, freshArrayBuffer } from "./fixtures";

export async function markdownBench(): Promise<Bench> {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));

  for (const fixture of FIXTURES) {
    const label = fixtureLabel(fixture);
    const doc = await folioParse(new Uint8Array(freshArrayBuffer(fixture)));
    const markdown = toMarkdown(doc);

    bench.add(`toMarkdown · ${label}`, () => {
      toMarkdown(doc);
    });
    bench.add(`fromMarkdown · ${label}`, () => {
      fromMarkdown(markdown);
    });
  }

  return bench;
}
