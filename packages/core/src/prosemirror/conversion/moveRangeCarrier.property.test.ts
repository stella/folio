/**
 * A tracked move keeps its name across the editor.
 *
 * `w:name` is the whole of a move's identity: `w:moveFromRangeStart` and
 * `w:moveToRangeStart` carry it, and it is the only thing that says this
 * deletion and that insertion are one relocation rather than two unrelated
 * revisions. The wrappers project as a tracked-change mark and the markers had
 * no node, so the editor round trip dropped the name, the range's `w:id` and
 * its author: a document saved after any edit had lost the move.
 *
 * The property generates the move's name, its ids and where the range sits, and
 * asserts the pairing survives `toProseDoc` → `fromProseDoc` and survives it
 * again once the changes are resolved.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import { parseDocx } from "../../docx/parser";
import { createEmptyDocx, repackDocx } from "../../docx/rezip";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const AUTHORED = 'w:author="A" w:date="2024-01-01T00:00:00Z"';

type MoveIds = { range: number; revision: number };

const movedAway = (name: string, { range, revision }: MoveIds): string =>
  `<w:p>` + movedAwayContent(name, { range, revision }) + `</w:p>`;

const movedAwayContent = (name: string, { range, revision }: MoveIds): string =>
  `<w:moveFromRangeStart w:id="${range}" w:name="${name}" ${AUTHORED}/>` +
  `<w:moveFrom w:id="${revision}" ${AUTHORED}>` +
  `<w:r><w:delText>relocated</w:delText></w:r></w:moveFrom>` +
  `<w:moveFromRangeEnd w:id="${range}"/>`;

const movedHere = (name: string, { range, revision }: MoveIds): string =>
  `<w:p>` +
  `<w:moveToRangeStart w:id="${range}" w:name="${name}" ${AUTHORED}/>` +
  `<w:moveTo w:id="${revision}" ${AUTHORED}>` +
  `<w:r><w:t>relocated</w:t></w:r></w:moveTo>` +
  `<w:moveToRangeEnd w:id="${range}"/>` +
  `</w:p>`;

const plainParagraph = `<w:p><w:r><w:t>untouched</w:t></w:r></w:p>`;

const documentWith = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentPartOf = async (saved: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(saved);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

/** The part as it comes back from one editor round trip. */
const throughTheEditor = async (body: string): Promise<string> => {
  const parsed = await parseDocx(await documentWith(body), { preloadFonts: false });
  const projected = fromProseDoc(toProseDoc(parsed), parsed);
  return documentPartOf(await repackDocx(projected, { updateModifiedDate: false }));
};

/** Every `w:name` the part carries, in document order. */
const namesIn = (xml: string): string[] =>
  [...xml.matchAll(/<w:(?:moveFrom|moveTo)RangeStart[^>]*\sw:name="([^"]*)"/gu)].map(
    // SAFETY: the capture group is present whenever the pattern matched.
    (match) => match[1] as string,
  );

/** A move name Word would accept: `ST_String`, non-empty, no XML metacharacters. */
const moveName = fc
  .stringMatching(/^[A-Za-z][A-Za-z0-9_]{0,16}$/u)
  .filter((value) => value.length > 0);

const annotationId = fc.integer({ min: 1, max: 5000 });

describe("a tracked move keeps its name across the editor", () => {
  test("both halves come back, paired by the name the document gave them", async () => {
    await fc.assert(
      fc.asyncProperty(
        moveName,
        fc.uniqueArray(annotationId, { minLength: 4, maxLength: 4 }),
        fc.boolean(),
        async (name, [fromRange, fromRevision, toRange, toRevision], separated) => {
          // SAFETY: `uniqueArray` of exactly four gives four defined ids.
          const from = { range: fromRange as number, revision: fromRevision as number };
          const to = { range: toRange as number, revision: toRevision as number };
          const body =
            movedAway(name, from) +
            (separated ? plainParagraph : "") +
            movedHere(name, to) +
            plainParagraph;

          const saved = await throughTheEditor(body);

          expect(namesIn(saved)).toEqual([name, name]);
          expect(saved).toContain(`<w:moveFromRangeEnd w:id="${from.range}"/>`);
          expect(saved).toContain(`<w:moveToRangeEnd w:id="${to.range}"/>`);
          // The range still opens before the content it delimits and closes
          // after it, which is the whole of what a delimiter pair says.
          expect(saved.indexOf("<w:moveFromRangeStart")).toBeLessThan(
            saved.indexOf("<w:moveFrom "),
          );
          expect(saved.indexOf("<w:moveFromRangeEnd")).toBeGreaterThan(
            saved.indexOf("<w:delText>relocated</w:delText>"),
          );
        },
      ),
      propertyConfig({ numRuns: 25 }),
    );
  });

  test("a second round trip is a fixed point", async () => {
    const body =
      movedAway("relocation", { range: 1, revision: 2 }) +
      movedHere("relocation", { range: 3, revision: 4 });
    const once = await throughTheEditor(body);
    const parsed = await parseDocx(await documentWith(body), { preloadFonts: false });
    const twice = await documentPartOf(
      await repackDocx(fromProseDoc(toProseDoc(fromProseDoc(toProseDoc(parsed), parsed)), parsed), {
        updateModifiedDate: false,
      }),
    );

    expect(namesIn(twice)).toEqual(namesIn(once));
  });

  test("two ranges of the same kind keep their separate extents", async () => {
    const body =
      `<w:p>` +
      movedAwayContent("first", { range: 11, revision: 21 }) +
      movedAwayContent("second", { range: 12, revision: 22 }) +
      `</w:p>`;

    const saved = await throughTheEditor(body);
    const boundaries = [
      saved.indexOf('<w:moveFromRangeStart w:id="11"'),
      saved.indexOf('<w:moveFrom w:id="21"'),
      saved.indexOf('<w:moveFromRangeEnd w:id="11"'),
      saved.indexOf('<w:moveFromRangeStart w:id="12"'),
      saved.indexOf('<w:moveFrom w:id="22"'),
      saved.indexOf('<w:moveFromRangeEnd w:id="12"'),
    ];

    expect(boundaries.every((position) => position >= 0)).toBe(true);
    expect(boundaries).toEqual(boundaries.toSorted((left, right) => left - right));
    expect(namesIn(saved)).toEqual(["first", "second"]);
  });

  test.each([
    ["bdo", '<w:bdo w:val="rtl">', "</w:bdo>"],
    ["dir", '<w:dir w:val="rtl">', "</w:dir>"],
    ["smartTag", '<w:smartTag w:element="Move">', "</w:smartTag>"],
    ["customXml", '<w:customXml w:element="Move">', "</w:customXml>"],
  ] as const)("keeps a non-empty move range inside w:%s", async (_kind, open, close) => {
    const saved = await throughTheEditor(
      `<w:p>${open}${movedAwayContent("nested", { range: 31, revision: 41 })}${close}</w:p>`,
    );

    expect(namesIn(saved)).toEqual(["nested"]);
    const rangeStart = saved.indexOf('<w:moveFromRangeStart w:id="31"');
    const move = saved.indexOf('<w:moveFrom w:id="41"');
    const rangeEnd = saved.indexOf('<w:moveFromRangeEnd w:id="31"');
    const startWrapper = saved.lastIndexOf(open, rangeStart);
    const startWrapperEnd = saved.indexOf(close, rangeStart);
    const moveWrapper = saved.indexOf(open, move);
    const moveWrapperEnd = saved.indexOf(close, moveWrapper);
    const endWrapper = saved.lastIndexOf(open, rangeEnd);
    const endWrapperEnd = saved.indexOf(close, rangeEnd);

    expect([rangeStart, move, rangeEnd].every((position) => position >= 0)).toBe(true);
    expect(startWrapper).toBeLessThan(rangeStart);
    expect(rangeStart).toBeLessThan(startWrapperEnd);
    expect(move).toBeLessThan(moveWrapper);
    expect(moveWrapper).toBeLessThan(moveWrapperEnd);
    expect(endWrapper).toBeLessThan(rangeEnd);
    expect(rangeEnd).toBeLessThan(endWrapperEnd);
  });
});
