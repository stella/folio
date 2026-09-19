/**
 * A range marker standing between two blocks must survive a save.
 *
 * `w:body`, `w:tc`, a header and an SDT's content all admit `EG_RunLevelElts`
 * and `EG_RangeMarkupElements` beside their paragraphs, and every block
 * container dropped them. `w:permStart` is the one that costs something a
 * reader can see: it is the whole of a document-protection range, so losing it
 * silently unprotects the saved file.
 *
 * Placement is the variable, not the marker: the same element is kept or lost
 * depending on whether it opens the body, separates two blocks, closes the
 * body, or sits in a table cell, so the property generates the position.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const AUTHORED = 'w:author="A" w:date="2024-01-01T00:00:00Z"';

/**
 * The markers a block container may hold, with the markup each one writes.
 *
 * `w:permStart` carries `w:edGrp`, the grant that makes a protected range mean
 * anything; the paired ends make each construct one Word would accept.
 */
const MARKERS = {
  perm: (id: number) => ({
    start: `<w:permStart w:id="${id}" w:edGrp="everyone"/>`,
    end: `<w:permEnd w:id="${id}"/>`,
  }),
  customXmlInsRange: (id: number) => ({
    start: `<w:customXmlInsRangeStart w:id="${id}" ${AUTHORED}/>`,
    end: `<w:customXmlInsRangeEnd w:id="${id}"/>`,
  }),
  customXmlMoveToRange: (id: number) => ({
    start: `<w:customXmlMoveToRangeStart w:id="${id}" ${AUTHORED}/>`,
    end: `<w:customXmlMoveToRangeEnd w:id="${id}"/>`,
  }),
  commentRange: (id: number) => ({
    start: `<w:commentRangeStart w:id="${id}"/>`,
    end: `<w:commentRangeEnd w:id="${id}"/>`,
  }),
  moveFromRange: (id: number) => ({
    start: `<w:moveFromRangeStart w:id="${id}" w:name="mv${id}" ${AUTHORED}/>`,
    end: `<w:moveFromRangeEnd w:id="${id}"/>`,
  }),
} as const;

type MarkerKind = keyof typeof MARKERS;

const MARKER_KINDS = Object.keys(MARKERS) as MarkerKind[];

const PLACEMENTS = {
  /** Before the first block, closing after it. */
  bodyStart: "bodyStart",
  /** Between two blocks, closing between the next two. */
  betweenBlocks: "betweenBlocks",
  /** After the last block, with nothing following it. */
  bodyEnd: "bodyEnd",
  /** Around a paragraph inside a table cell. */
  inCell: "inCell",
} as const;

type Placement = (typeof PLACEMENTS)[keyof typeof PLACEMENTS];

const PLACEMENT_VALUES = Object.values(PLACEMENTS);

const paragraph = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const table = (inner: string): string =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr/>${inner}</w:tc></w:tr></w:tbl>`;

const bodyFor = (placement: Placement, { start, end }: { start: string; end: string }): string => {
  switch (placement) {
    case PLACEMENTS.bodyStart:
      return `${start}${paragraph("one")}${end}${paragraph("two")}`;
    case PLACEMENTS.betweenBlocks:
      return `${paragraph("one")}${start}${paragraph("two")}${end}${paragraph("three")}`;
    case PLACEMENTS.bodyEnd:
      return `${paragraph("one")}${start}${paragraph("two")}${end}`;
    case PLACEMENTS.inCell:
      return `${paragraph("one")}${table(`${start}${paragraph("cell")}${end}`)}`;
    default: {
      const unreachable: never = placement;
      return unreachable;
    }
  }
};

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const savedDocumentXml = async (body: string): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const saved = await repackDocx(parsed, { updateModifiedDate: false });
  const zip = await JSZip.loadAsync(saved);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

/** Where each block-level marker sits relative to the text around it. */
const positionsOf = (xml: string, markers: readonly string[]): number[] =>
  markers.map((marker) => xml.indexOf(marker));

describe("a range marker between blocks survives a save", () => {
  test("every marker kind, in every block position, comes back in place", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...MARKER_KINDS),
        fc.constantFrom(...PLACEMENT_VALUES),
        fc.integer({ min: 1, max: 5000 }),
        async (kind, placement, id) => {
          const marker = MARKERS[kind](id);
          const saved = await savedDocumentXml(bodyFor(placement, marker));

          const [startAt, endAt] = positionsOf(saved, [marker.start, marker.end]);
          expect({ kind, placement, start: startAt !== -1, end: endAt !== -1 }).toEqual({
            kind,
            placement,
            start: true,
            end: true,
          });
          // The range still opens before it closes, which is the whole of what
          // a delimiter pair says.
          expect(startAt).toBeLessThan(endAt ?? -1);
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("a protected range keeps its grant and its place around the paragraph it protects", async () => {
    const marker = MARKERS.perm(900);
    const saved = await savedDocumentXml(bodyFor(PLACEMENTS.betweenBlocks, marker));

    expect(saved).toContain('<w:permStart w:id="900" w:edGrp="everyone"/>');
    expect(saved).toContain('<w:permEnd w:id="900"/>');
    const protectedParagraph = saved.indexOf("<w:t>two</w:t>");
    expect(saved.indexOf(marker.start)).toBeLessThan(protectedParagraph);
    expect(saved.indexOf(marker.end)).toBeGreaterThan(protectedParagraph);
  });
});
