/**
 * Serialize benchmark: document model → DOCX bytes.
 *
 * Each fixture is parsed once during setup, then only its serialization is
 * measured.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import { createDocx as folioCreate } from "@stll/folio-core";
import { parseDocx as folioParse } from "@stll/folio-core/docx/parser";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";
import { FIXTURES, fixtureLabel, freshArrayBuffer } from "./fixtures";

export async function serializeBench(): Promise<Bench> {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));

  for (const fixture of FIXTURES) {
    const label = fixtureLabel(fixture);

    const folioDoc = await folioParse(new Uint8Array(freshArrayBuffer(fixture)));
    bench.add(`folio · ${label}`, async () => {
      await folioCreate(folioDoc);
    });
  }

  return bench;
}
