import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { MAX_REVISION_ID } from "@stll/docx-core/model";
import { createNumericIdAllocator } from "./numericIdAllocator";
import { createNumberingIdAllocator, mintNumberingId } from "./numberingIds";
import { reserveBookmarkIds, mintBookmarkId } from "./bookmarkIds";
import { endnote } from "./server/build";
import { createEmptyDocument } from "../utils/createDocument";

setDefaultTimeout(propertyTestTimeout(10_000));

test("reservation and rollover preserve range and uniqueness over allocation sequences", () => {
  fc.assert(
    fc.property(
      fc.array(fc.integer({ min: 0, max: MAX_REVISION_ID }), { maxLength: 30 }),
      fc.constantFrom(0, 1),
      (existingIds, firstId) => {
        const allocator = createNumericIdAllocator({ space: "test", firstId });
        const existing = [
          firstId,
          firstId + 1,
          MAX_REVISION_ID - 1,
          MAX_REVISION_ID,
          ...existingIds,
        ];
        allocator.reserve(existing);
        const used = new Set(existing);
        for (let index = 0; index < 40; index++) {
          const id = allocator.next();
          expect(id).toBeGreaterThanOrEqual(firstId);
          expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
          expect(used.has(id)).toBe(false);
          used.add(id);
          // Repeated loads cannot free an already minted id during rollover.
          allocator.reserve(existing);
        }
      },
    ),
    propertyConfig({ numRuns: 30 }),
  );
});

test("bookmark and numbering spaces skip reserved ids after a loaded maximum", () => {
  reserveBookmarkIds([0, 1, MAX_REVISION_ID]);
  const bookmarkId = mintBookmarkId();
  expect(bookmarkId).toBeGreaterThan(1);
  expect(bookmarkId).toBeLessThan(MAX_REVISION_ID);
  for (const kind of ["num", "abstract"] as const) {
    const id = mintNumberingId({ kind, existingIds: [0, 1, MAX_REVISION_ID] });
    expect(id).toBeGreaterThan(1);
    expect(id).toBeLessThan(MAX_REVISION_ID);
  }
});

test("offline numbering allocators preserve deterministic independent package output", () => {
  const existingIds = [0, 1, MAX_REVISION_ID];
  const first = createNumberingIdAllocator("abstract", existingIds);
  const second = createNumberingIdAllocator("abstract", existingIds);
  expect([first.next(), first.next()]).toEqual([second.next(), second.next()]);
});

test("numbering allocation depends only on the document's reserved ids", () => {
  for (const kind of ["num", "abstract"] as const) {
    const existingIds = [0, 1, MAX_REVISION_ID];
    const first = mintNumberingId({ kind, existingIds });
    mintNumberingId({ kind, existingIds: [0, 1, 2, 3, 4] });
    expect(mintNumberingId({ kind, existingIds })).toBe(first);
    expect(mintNumberingId({ kind, existingIds: [...existingIds, first] })).not.toBe(first);
  }
});

test("the endnote builder allocates a free bounded id when loaded notes contain the maximum", () => {
  const doc = createEmptyDocument();
  doc.package.endnotes = [
    { type: "endnote", id: 1, content: [] },
    { type: "endnote", id: MAX_REVISION_ID, content: [] },
  ];
  const reference = endnote(doc, "Note");
  const note = doc.package.endnotes.at(-1);
  expect(note?.id).toBeGreaterThan(1);
  expect(note?.id).toBeLessThan(MAX_REVISION_ID);
  expect(reference.content.at(0)).toEqual({ type: "endnoteRef", id: note?.id });
});
