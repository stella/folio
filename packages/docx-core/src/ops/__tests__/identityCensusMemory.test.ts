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

/**
 * Live `Map` objects once collection settles. What earlier work left can
 * take more than one collection to free (a weak entry frees its value only
 * once its key is collected), so collect until the count holds.
 */
const liveMaps = (): number => {
  let count = Number.NaN;
  for (let round = 0; round < 16; round += 1) {
    Bun.gc(true);
    const now = heapStats().objectTypeCounts["Map"] ?? 0;
    if (now === count) return now;
    count = now;
  }
  return count;
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
  // The indexes themselves hold a few maps of the keys they count, however
  // many records the package has; the memo of what each record contributes
  // adds none per record, where one per record would be thousands here.
  for (const paragraphs of [20, 2_000])
    expect(mapsHeldByCensus(packageOf(paragraphs))).toBeLessThan(20);
});
