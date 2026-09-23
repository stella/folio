/**
 * ProseMirror bridge benchmark: document model ↔ ProseMirror document.
 *
 * Every open converts the parsed model into a ProseMirror document, and every
 * save converts it back, so both directions sit on the interactive path. Each
 * fixture is parsed and converted once during setup; only the conversion
 * itself is measured.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import { parseDocx as folioParse } from "@stll/folio-core/docx/parser";
import { fromProseDoc, toProseDoc } from "@stll/folio-core/prosemirror/conversion";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";
import { FIXTURES, fixtureLabel, freshArrayBuffer } from "./fixtures";

export async function proseMirrorBench(): Promise<Bench> {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));

  for (const fixture of FIXTURES) {
    const label = fixtureLabel(fixture);
    const doc = await folioParse(new Uint8Array(freshArrayBuffer(fixture)));
    const proseDoc = toProseDoc(doc);

    bench.add(`toProseDoc · ${label}`, () => {
      toProseDoc(doc);
    });
    bench.add(`fromProseDoc · ${label}`, () => {
      fromProseDoc(proseDoc, doc);
    });
  }

  return bench;
}
