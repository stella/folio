/**
 * A fixed package whose parse time is a reading of the machine.
 *
 * The corpus gate prices every file against the corpus, which takes the file's
 * size out of the verdict. It does not take the machine out: a run is minutes
 * long and the load on a shared box is not, so a file parsed during a spike is
 * compared against a baseline set mostly by files parsed while things were
 * quiet. The same file then passes or fails between identical runs.
 *
 * This package never changes, so what it costs is the machine and nothing
 * else. Parsing it beside each corpus file gives that file a reading of the
 * load it was measured under, and the performance family divides it out.
 *
 * It is deliberately small. The reading has to be cheap enough to take
 * thousands of times without changing what the gate costs, and it only has to
 * resolve the load, not characterise the parser.
 */

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

/**
 * Built once per worker process, not per file: building it is not the work
 * being measured, and rebuilding it would put the serializer's cost into a
 * reading meant to hold only the parser's.
 */
let referencePackage: ArrayBuffer | null = null;

const buildReferencePackage = (): Promise<ArrayBuffer> => createDocx(createEmptyDocument());

/**
 * Milliseconds the reference package took to parse, now.
 *
 * Zero when the reference itself could not be built or parsed, which the
 * performance family reads as "this file was never priced against the
 * machine" and declines a verdict on rather than guessing.
 */
export const measureReferenceMs = async (): Promise<number> => {
  try {
    const first = referencePackage === null;
    referencePackage ??= await buildReferencePackage();
    if (first) {
      // The very first parse in a process pays for compiling the parser, which
      // is the process starting rather than the machine being busy. Charging
      // it to the first file would scale that file's cost down against a
      // reference it never really ran under, so it is spent here instead.
      await parseDocx(referencePackage, { preloadFonts: false });
    }
    const started = Bun.nanoseconds();
    await parseDocx(referencePackage, { preloadFonts: false });
    return (Bun.nanoseconds() - started) / 1e6;
  } catch {
    // A reference that cannot be measured must not fail the file being
    // measured; the family already knows what a missing reading means.
    return 0;
  }
};
