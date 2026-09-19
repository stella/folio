/**
 * A revision id folio mints must not land on another annotation's id.
 *
 * `w:id` on a comment, a bookmark, a protected range and a tracked change all
 * come from one counter in Word: of the corpus packages carrying more than one
 * of those kinds, almost none repeats a value across them, which separate
 * counters starting at zero could not produce. `commentIdAllocator.ts` states
 * the same rule. The save-time deduplication allocated the lowest free integer
 * while reserving only revision ids, so renumbering a duplicated `w:ins` in a
 * commented document could hand it a live comment's id.
 *
 * The property generates the collision rather than an example of it: which id
 * is free depends on every id in the package, so the arrangement is the
 * variable.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { normalizeRevisionIdsInXmlParts } from "./revisionIdNormalization";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Small ids, so the allocator's lowest-free search actually has to step over them. */
const annotationId = fc.integer({ min: 0, max: 12 });

type GeneratedPackage = {
  commentIds: number[];
  bookmarkIds: number[];
  permIds: number[];
  moveRangeIds: number[];
  /** Revision ids in document order; a repeat is what forces a rewrite. */
  revisionIds: number[];
};

const generatedPackage = fc.record({
  commentIds: fc.uniqueArray(annotationId, { maxLength: 4 }),
  bookmarkIds: fc.uniqueArray(annotationId, { maxLength: 4 }),
  permIds: fc.uniqueArray(annotationId, { maxLength: 3 }),
  moveRangeIds: fc.uniqueArray(annotationId, { maxLength: 3 }),
  revisionIds: fc.array(annotationId, { minLength: 1, maxLength: 6 }),
});

const documentXml = ({
  commentIds,
  bookmarkIds,
  permIds,
  moveRangeIds,
  revisionIds,
}: GeneratedPackage): string =>
  `<w:document ${W}><w:body><w:p>` +
  commentIds
    .map(
      (id) =>
        `<w:commentRangeStart w:id="${id}"/><w:r><w:t>c</w:t></w:r>` +
        `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`,
    )
    .join("") +
  bookmarkIds
    .map((id) => `<w:bookmarkStart w:id="${id}" w:name="bm${id}"/><w:bookmarkEnd w:id="${id}"/>`)
    .join("") +
  permIds.map((id) => `<w:permStart w:id="${id}"/><w:permEnd w:id="${id}"/>`).join("") +
  moveRangeIds
    .map(
      (id) =>
        `<w:moveFromRangeStart w:id="${id}" w:name="mv${id}" w:author="A"/>` +
        `<w:moveFromRangeEnd w:id="${id}"/>`,
    )
    .join("") +
  revisionIds
    .map((id) => `<w:ins w:id="${id}" w:author="A"><w:r><w:t>i</w:t></w:r></w:ins>`)
    .join("") +
  `</w:p></w:body></w:document>`;

const commentsXml = (commentIds: readonly number[]): string =>
  `<w:comments ${W}>` +
  commentIds
    .map(
      (id) => `<w:comment w:id="${id}" w:author="A"><w:p><w:r><w:t>n</w:t></w:r></w:p></w:comment>`,
    )
    .join("") +
  `</w:comments>`;

const buildParts = (generated: GeneratedPackage): Map<string, string> =>
  new Map([
    ["word/document.xml", documentXml(generated)],
    ["word/comments.xml", commentsXml(generated.commentIds)],
  ]);

const numbersIn = (xml: string, element: string): number[] =>
  [...xml.matchAll(new RegExp(`<w:${element}\\b[^>]*\\bw:id="(\\d+)"`, "gu"))].map(([, value]) =>
    Number(value),
  );

const revisionIdsOf = (xml: string): number[] => numbersIn(xml, "ins");

const otherAnnotationIdsOf = (xml: string): Set<number> =>
  new Set([
    ...numbersIn(xml, "commentRangeStart"),
    ...numbersIn(xml, "commentRangeEnd"),
    ...numbersIn(xml, "commentReference"),
    ...numbersIn(xml, "bookmarkStart"),
    ...numbersIn(xml, "bookmarkEnd"),
    ...numbersIn(xml, "permStart"),
    ...numbersIn(xml, "permEnd"),
    ...numbersIn(xml, "moveFromRangeStart"),
    ...numbersIn(xml, "moveFromRangeEnd"),
  ]);

describe("revision ids are minted out of the whole annotation space", () => {
  test("a minted id never lands on a comment, bookmark, perm or move-range id", () => {
    fc.assert(
      fc.property(generatedPackage, (generated) => {
        const parts = buildParts(generated);
        const source = parts.get("word/document.xml") ?? "";
        const normalized = normalizeRevisionIdsInXmlParts(parts);
        const document = normalized.get("word/document.xml") ?? "";
        const taken = otherAnnotationIdsOf(source);
        const before = revisionIdsOf(source);
        // Only the ids this pass chose. An input that already reuses a value
        // across kinds is the producer's, and copying it through is faithful.
        for (const [index, id] of revisionIdsOf(document).entries()) {
          if (id === before[index]) {
            continue;
          }
          expect({ id, taken: taken.has(id) }).toEqual({ id, taken: false });
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("revision ids are unique, and an already-unique set is untouched", () => {
    fc.assert(
      fc.property(generatedPackage, (generated) => {
        const parts = buildParts(generated);
        const normalized = normalizeRevisionIdsInXmlParts(parts);
        const document = normalized.get("word/document.xml") ?? "";
        const ids = revisionIdsOf(document);
        expect(new Set(ids).size).toBe(ids.length);

        // Nothing is rewritten when the revision ids were already unique.
        if (new Set(generated.revisionIds).size === generated.revisionIds.length) {
          expect(ids).toEqual(generated.revisionIds);
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("the pass is a fixed point", () => {
    fc.assert(
      fc.property(generatedPackage, (generated) => {
        const once = normalizeRevisionIdsInXmlParts(buildParts(generated));
        expect(normalizeRevisionIdsInXmlParts(once)).toEqual(once);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
