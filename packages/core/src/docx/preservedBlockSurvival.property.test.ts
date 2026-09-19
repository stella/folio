/**
 * A block child folio does not model survives the editor, in place.
 *
 * The block containers modelled paragraphs, tables and content controls and
 * let everything else fall off the end of the walk: a `w:permStart` standing
 * between two paragraphs is the whole of a document-protection range, a
 * `w:altChunk` is an entire imported document, an `m:oMathPara` is a display
 * equation. Position is their meaning, so the property generates the position
 * and asserts both halves — the markup comes back, and it comes back between
 * the same two blocks.
 *
 * The editor leg is the one that used to fail even when the save did not: the
 * capture rode on the neighbouring block and `toProseDoc` had nowhere to put
 * it, so a document that survived an untouched save lost the markup the moment
 * anybody opened it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/**
 * One captured child per kind folio has no model for, spanning the reasons a
 * container may hold one: a protection range, an imported document, a display
 * equation, a proofing hint and a custom-XML wrapper.
 */
const PRESERVED = {
  permStart: '<w:permStart w:id="77" w:edGrp="everyone"/>',
  altChunk: '<w:altChunk r:id="rIdChunk"/>',
  oMathPara: `<m:oMathPara xmlns:m="${M_NAMESPACE}"><m:oMath/></m:oMathPara>`,
  proofErr: '<w:proofErr w:type="spellStart"/>',
  customXml: '<w:customXml w:element="clause"><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:customXml>',
} as const;

type PreservedKind = keyof typeof PRESERVED;

const PRESERVED_KINDS = Object.keys(PRESERVED) as PreservedKind[];

const paragraph = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const table = (text: string): string =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr/>${paragraph(text)}</w:tc></w:tr></w:tbl>`;

/** Where the capture sits relative to the blocks folio does model. */
const PLACEMENTS = {
  /** Before every block, with nothing in front of it. */
  first: "first",
  /** Between two paragraphs. */
  betweenParagraphs: "betweenParagraphs",
  /** Between a paragraph and a table. */
  beforeTable: "beforeTable",
  /** After every block, with nothing behind it. */
  last: "last",
  /** Alone, in a container holding nothing else. */
  alone: "alone",
} as const;

type Placement = (typeof PLACEMENTS)[keyof typeof PLACEMENTS];

const PLACEMENT_VALUES = Object.values(PLACEMENTS);

const bodyFor = (placement: Placement, captured: string): string => {
  switch (placement) {
    case PLACEMENTS.first:
      return `${captured}${paragraph("one")}${paragraph("two")}`;
    case PLACEMENTS.betweenParagraphs:
      return `${paragraph("one")}${captured}${paragraph("two")}`;
    case PLACEMENTS.beforeTable:
      return `${paragraph("one")}${captured}${table("cell")}`;
    case PLACEMENTS.last:
      return `${paragraph("one")}${paragraph("two")}${captured}`;
    case PLACEMENTS.alone:
      return captured;
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
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXmlOf = async (saved: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";

/** Parse, project through the editor and back, save: the trip that used to lose it. */
const throughTheEditor = async (body: string): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const projected = fromProseDoc(toProseDoc(parsed), parsed);
  return documentXmlOf(await repackDocx(projected, { updateModifiedDate: false }));
};

/** Parse and save, with no editor in between. */
const throughASave = async (body: string): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  return documentXmlOf(await repackDocx(parsed, { updateModifiedDate: false }));
};

/**
 * The marker a capture leaves in the saved part.
 *
 * A capture is re-spelled on the way out — attribute order, an empty element
 * written long — so the assertion is on the element name rather than on the
 * authored string.
 */
const markerFor = (kind: PreservedKind): string =>
  kind === "oMathPara" ? "<m:oMathPara" : `<w:${kind}`;

describe("a block child folio does not model survives in place", () => {
  test("every kind, in every position, comes back between the same blocks", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...PRESERVED_KINDS),
        fc.constantFrom(...PLACEMENT_VALUES),
        async (kind, placement) => {
          const saved = await throughTheEditor(bodyFor(placement, PRESERVED[kind]));
          const at = saved.indexOf(markerFor(kind));
          expect({ kind, placement, kept: at !== -1 }).toEqual({ kind, placement, kept: true });

          const one = saved.indexOf("<w:t>one</w:t>");
          const second = saved.indexOf("<w:t>two</w:t>");
          const cell = saved.indexOf("<w:t>cell</w:t>");
          switch (placement) {
            case PLACEMENTS.first:
              expect(at).toBeLessThan(one);
              break;
            case PLACEMENTS.betweenParagraphs:
              expect(at).toBeGreaterThan(one);
              expect(at).toBeLessThan(second);
              break;
            case PLACEMENTS.beforeTable:
              expect(at).toBeGreaterThan(one);
              expect(at).toBeLessThan(cell);
              break;
            case PLACEMENTS.last:
              expect(at).toBeGreaterThan(second);
              break;
            case PLACEMENTS.alone:
              break;
            default: {
              const unreachable: never = placement;
              return unreachable;
            }
          }
          return true;
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("a save with no editor in between keeps it too", async () => {
    const saved = await throughASave(bodyFor(PLACEMENTS.betweenParagraphs, PRESERVED.permStart));
    expect(saved).toContain('<w:permStart w:id="77" w:edGrp="everyone"/>');
    expect(saved.indexOf("<w:permStart")).toBeGreaterThan(saved.indexOf("<w:t>one</w:t>"));
    expect(saved.indexOf("<w:permStart")).toBeLessThan(saved.indexOf("<w:t>two</w:t>"));
  });

  test("a header keeps one, and so does a table cell", async () => {
    const cellBody =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr/>${paragraph("before")}${PRESERVED.permStart}` +
      `${paragraph("after")}</w:tc></w:tr></w:tbl>`;
    const saved = await throughTheEditor(cellBody);
    expect(saved).toContain("<w:permStart");
    expect(saved.indexOf("<w:permStart")).toBeGreaterThan(saved.indexOf("<w:t>before</w:t>"));
    expect(saved.indexOf("<w:permStart")).toBeLessThan(saved.indexOf("<w:t>after</w:t>"));
  });
});
