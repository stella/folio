import { heapStats } from "bun:jsc";
import { expect, test } from "bun:test";

import type { DocxPackage, Paragraph } from "../../model/document";
import { formatParaId } from "../idLease";
import {
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  packageReservedKeyIndex,
} from "../ids";

const packageOf = (paragraphs: number): DocxPackage => ({
  document: {
    content: Array.from({ length: paragraphs }, (_, index): Paragraph => ({
      type: "paragraph",
      paraId: formatParaId(index + 1),
      content: [
        {
          type: "insertion",
          info: { id: index + 1, author: "Editor" },
          content: [{ type: "run", content: [{ type: "text", text: "Text" }] }],
        },
      ],
    })),
  },
});

/** Live `Map` objects once the collector has run. */
const liveMaps = (): number => {
  Bun.gc(true);
  return heapStats().objectTypeCounts.Map ?? 0;
};

/** How many more maps stay live while a census of `pkg` is remembered. */
const mapsHeldByCensus = (pkg: DocxPackage): number => {
  const before = liveMaps();
  packageParagraphIdIndex(pkg);
  packageIdentityKeyIndex(pkg);
  packageReservedKeyIndex(pkg);
  const after = liveMaps();
  // Keep the package, and so what its census remembers, live until measured.
  expect(pkg.document.content.length).toBeGreaterThan(0);
  return after - before;
};

test("what a package census remembers does not grow a map per record", () => {
  const small = mapsHeldByCensus(packageOf(20));
  const large = mapsHeldByCensus(packageOf(2_000));
  // The indexes themselves hold maps of the keys they count; the memo of
  // what each record contributes adds none per record.
  expect(large - small).toBeLessThan(20);
});
