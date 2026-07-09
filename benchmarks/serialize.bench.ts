/**
 * Serialize benchmark: document model → DOCX bytes.
 *
 * folio vs `@eigenpal/docx-editor-core` (the upstream). Each library parses the
 * fixture into ITS OWN model once (setup, not timed) — the models are different
 * types — then we bench only its serialize step (`createDocx`).
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import {
  createDocx as eigenpalCreate,
  parseDocx as eigenpalParse,
} from "@eigenpal/docx-editor-core";
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

    const eigenpalDoc = await eigenpalParse(new Uint8Array(freshArrayBuffer(fixture)));
    bench.add(`eigenpal · ${label}`, async () => {
      await eigenpalCreate(eigenpalDoc);
    });
  }

  return bench;
}
