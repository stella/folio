/**
 * A rebuilt drawing keeps the name and alt text it was authored with.
 *
 * `pic:cNvPr` and `wps:cNvPr` are the picture's and the shape's own
 * `CT_NonVisualDrawingProps`: `@name` is the accessible name a reader announces
 * and `@descr` is the alt text. The rebuild spelled `pic:cNvPr@name` from the
 * media filename and `@descr` from the drawing's alt text, so resizing a
 * picture renamed it `image1.png` and moved the drawing's alt text onto it, and
 * it wrote no `wps:cNvPr` at all, so a shape named inside its graphic lost that
 * name on every edit. Both are accessibility regressions rather than
 * formatting ones.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

const IMAGE_RID = "rIdPicture";

/** A 1×1 PNG, so the drawing has a real picture relationship to resolve. */
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);

/**
 * The inner name set, distinct from `wp:docPr`'s so a writer that copies one
 * onto the other cannot pass: `@name` is a caption, `@descr` alt text.
 */
const INNER = { name: "Company seal", descr: "the seal of the issuing office" } as const;

const NAMES_XML = `name="${INNER.name}" descr="${INNER.descr}"`;

const PICTURE_GRAPHIC =
  `<a:graphic><a:graphicData uri="${PIC_NAMESPACE}">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" ${NAMES_XML}/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${IMAGE_RID}"/></pic:blipFill><pic:spPr/></pic:pic>` +
  "</a:graphicData></a:graphic>";

const SHAPE_GRAPHIC =
  `<a:graphic><a:graphicData uri="${WPS_NAMESPACE}">` +
  `<wps:wsp><wps:cNvPr id="2" ${NAMES_XML}/><wps:cNvSpPr/>` +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>";

const GRAPHIC_BY_KIND = { picture: PICTURE_GRAPHIC, shape: SHAPE_GRAPHIC } as const;

type DrawingKind = keyof typeof GRAPHIC_BY_KIND;

const DRAWING_KINDS = Object.keys(GRAPHIC_BY_KIND) as DrawingKind[];

/** The drawing's own names, which the inner ones must not be confused with. */
const DOC_PR = '<wp:docPr id="1" name="Drawing 1" descr="the drawing, not the picture"/>';

const bodyFor = (kind: DrawingKind): string =>
  "<w:p><w:r><w:drawing><wp:inline>" +
  '<wp:extent cx="914400" cy="914400"/>' +
  `${DOC_PR}${GRAPHIC_BY_KIND[kind]}</wp:inline></w:drawing></w:r></w:p>`;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}" ` +
      `xmlns:a="${A_NAMESPACE}" xmlns:wp="${WP_NAMESPACE}" xmlns:pic="${PIC_NAMESPACE}" ` +
      `xmlns:wps="${WPS_NAMESPACE}">` +
      `<w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  zip.file("word/media/image1.png", PNG);
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  zip.file(
    "[Content_Types].xml",
    types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="${IMAGE_RID}" Type="${R_NAMESPACE}/image" Target="media/image1.png"/>` +
        "</Relationships>",
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Resize every drawing, which is what stops the captured bytes being replayed. */
const resizeDrawings = (document: Document): void => {
  const visit = (value: unknown, seen: WeakSet<object>): void => {
    if (Array.isArray(value)) {
      for (const item of value) {
        visit(item, seen);
      }
      return;
    }
    if (typeof value !== "object" || value === null || seen.has(value)) {
      return;
    }
    seen.add(value);
    const record = value as Record<string, unknown>;
    if (
      (record["type"] === "image" || record["type"] === "shape") &&
      typeof record["size"] === "object"
    ) {
      record["size"] = { width: 457_200, height: 457_200 };
    }
    for (const item of Object.values(record)) {
      visit(item, seen);
    }
  };
  visit(document.package, new WeakSet());
};

const savedAfterResize = async (kind: DrawingKind): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(bodyFor(kind)), { preloadFonts: false });
  resizeDrawings(parsed);
  const saved = await repackDocx(parsed, { updateModifiedDate: false });
  return (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
};

describe("a rebuilt drawing keeps its authored non-visual names", () => {
  test("the picture's and the shape's own name and alt text survive a resize", async () => {
    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...DRAWING_KINDS), async (kind) => {
        const saved = await savedAfterResize(kind);
        expect({ kind, name: saved.includes(`name="${INNER.name}"`) }).toEqual({
          kind,
          name: true,
        });
        expect({ kind, descr: saved.includes(`descr="${INNER.descr}"`) }).toEqual({
          kind,
          descr: true,
        });
        // The media filename is not a name anybody wrote, so a drawing that
        // was named does not get it.
        expect(saved).not.toContain('name="image1.png"');
        // The drawing's own names stay on `wp:docPr` and reach neither inner
        // element.
        expect(saved).toContain('name="Drawing 1"');
      }),
      propertyConfig({ numRuns: DRAWING_KINDS.length }),
    );
  });
});
