/**
 * A preview is a render of markup folio cannot write, so an edit to it can
 * only be refused.
 *
 * Folio draws a preview for a drawing it cannot project: a VML shape, whose
 * markup it has no serializer for at all, and a diagram, which it describes
 * rather than draws. The preview is the picture the editor shows, and the
 * captured markup is the drawing. Editing the picture cannot describe the
 * drawing, so the only save that keeps the document is the one that replays
 * the capture and reports the edit as lost.
 *
 * The VML preview did neither. It named its relationship `""`, which is a key
 * no relationship answers to but is not the model's spelling for absence, and
 * it carried no mode, so the editor treated the render as an ordinary editable
 * picture: the first resize dropped the capture, and the save wrote folio's own
 * SVG into `word/media/` as the picture the shape had become. With `src` past
 * the preview budget the same save wrote `<a:blip r:embed=""/>` instead.
 *
 * The property runs every edit the commit seam can make over both preview
 * kinds and demands the package back: the authored markup verbatim, no media
 * part folio invented, no relationship it minted, and the same drawing on
 * reopen.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, DrawingContent } from "../types/document";
import { expectImageAttrs, mergeImageAttrs } from "../prosemirror/attrs";
import { allowsDirectDrawingEdit, classifyDrawingSafety } from "./imageRawXml";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { ImageAttrs } from "../prosemirror/schema/nodes";
import { parseDocx, parseDocxWithPreviewBudget } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const DOCUMENT_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
].join(" ");

/** A VML rectangle: artwork with no picture behind it, so folio renders it. */
const VML_SHAPE_MARKUP =
  '<w:pict><v:rect id="Rect 1" style="width:120pt;height:40pt" fillcolor="#ff0000" strokecolor="#123456"/></w:pict>';

/** A SmartArt frame: no `a:blip` anywhere, so folio describes it instead. */
const DIAGRAM_MARKUP =
  '<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="4" name="Diagram 4"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic></wp:inline></w:drawing>';

const DIAGRAM_DATA_XML = `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rIdDrawing"/></a:ext></dgm:extLst></dgm:dataModel>`;

const DIAGRAM_DRAWING_XML = `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="10000" y="10000"/><a:ext cx="400000" cy="200000"/></a:xfrm><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`;

const PREVIEW_KINDS_UNDER_TEST = ["vmlShape", "diagram"] as const;
type PreviewKindUnderTest = (typeof PREVIEW_KINDS_UNDER_TEST)[number];

const AUTHORED_MARKUP = {
  vmlShape: VML_SHAPE_MARKUP,
  diagram: DIAGRAM_MARKUP,
} as const satisfies Record<PreviewKindUnderTest, string>;

/** The markup a reader can find the drawing by, whatever else a save rewrites. */
const AUTHORED_SIGNATURE = {
  vmlShape: 'fillcolor="#ff0000" strokecolor="#123456"',
  diagram: '<dgm:relIds r:dm="rIdData"/>',
} as const satisfies Record<PreviewKindUnderTest, string>;

const diagramRelationships = (kind: PreviewKindUnderTest): string =>
  kind === "diagram"
    ? `<Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>`
    : "";

const packageWith = (kind: PreviewKindUnderTest): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${diagramRelationships(kind)}</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${DOCUMENT_NAMESPACES}><w:body><w:p><w:r>${AUTHORED_MARKUP[kind]}</w:r></w:p></w:body></w:document>`,
  );
  if (kind === "diagram") {
    zip.file("word/diagrams/data1.xml", DIAGRAM_DATA_XML);
    zip.file("word/diagrams/drawing1.xml", DIAGRAM_DRAWING_XML);
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Every edit the image commit seam can make to a drawing's projection. */
const EDITS = {
  resize: { width: 300, height: 100 },
  floatMove: {
    position: {
      horizontal: { posOffset: 123_456, relativeTo: "margin" },
      vertical: { posOffset: 654_321, relativeTo: "margin" },
    },
  },
  altText: { alt: "A rectangle" },
  wrapChange: { wrapType: "square" },
} as const satisfies Record<string, Partial<ImageAttrs>>;

type EditName = keyof typeof EDITS;
const EDIT_NAMES = Object.keys(EDITS) as EditName[];

const drawingOf = (document: Document): DrawingContent => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("the fixture lost its paragraph");
  }
  const run = paragraph.content.at(0);
  if (run?.type !== "run") {
    throw new Error("the fixture lost its run");
  }
  const drawing = run.content.at(0);
  if (drawing?.type !== "drawing") {
    throw new Error("the fixture lost its drawing");
  }
  return drawing;
};

const imageNodeIn = (document: PMNode): { node: PMNode; position: number } => {
  let found: { node: PMNode; position: number } | undefined;
  document.descendants((node, position) => {
    if (node.type.name === "image") {
      found = { node, position };
    }
    return true;
  });
  if (!found) {
    throw new Error("the projection lost its image node");
  }
  return found;
};

/**
 * Apply an edit the way a host that ignores the refusal would: straight onto
 * the node's attrs, through the merge the commit seam uses.
 */
const withEditedImage = (projected: PMNode, patch: Partial<ImageAttrs>): PMNode => {
  const { node } = imageNodeIn(projected);
  const edited = node.type.create(mergeImageAttrs(node, patch));
  const paragraph = projected.child(0);
  return projected.copy(
    projected.content.replaceChild(0, paragraph.copy(paragraph.content.replaceChild(0, edited))),
  );
};

type SavedPackage = {
  documentXml: string;
  relationshipsXml: string;
  entries: readonly string[];
};

const savedPackageOf = async (document: Document): Promise<SavedPackage> => {
  const zip = await JSZip.loadAsync(await repackDocx(document, { updateModifiedDate: false }));
  return {
    documentXml: (await zip.file("word/document.xml")?.async("string")) ?? "",
    relationshipsXml: (await zip.file("word/_rels/document.xml.rels")?.async("string")) ?? "",
    entries: Object.keys(zip.files),
  };
};

describe("a preview names no relationship", () => {
  test.each(PREVIEW_KINDS_UNDER_TEST)("%s", async (kind) => {
    const document = await parseDocx(await packageWith(kind), { preloadFonts: false });
    const drawing = drawingOf(document);

    expect(drawing.image.rId).toBeUndefined();
    // And the drawing says the picture is a render, so the editor declines it
    // rather than letting an edit reach a model that cannot describe it.
    expect(drawing.rawXmlMode).toBeDefined();
    expect(allowsDirectDrawingEdit(drawing.rawXmlMode)).toBe(false);
  });
});

describe("an edited preview saves the drawing it was made from", () => {
  test("every edit, every preview kind, with and without the render retained", async () => {
    const sources = new Map(
      await Promise.all(
        PREVIEW_KINDS_UNDER_TEST.map(
          async (kind) => [kind, await packageWith(kind)] as const satisfies [string, ArrayBuffer],
        ),
      ),
    );

    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...PREVIEW_KINDS_UNDER_TEST),
        fc.constantFrom(...EDIT_NAMES),
        // The render is what a package over the preview budget loses. A drawing
        // whose `src` is gone still reserves its space on the page, and it is
        // the same drawing to a save.
        fc.boolean(),
        async (kind, edit, starvePreview) => {
          const source = sources.get(kind);
          if (source === undefined) {
            throw new Error(`no package for ${kind}`);
          }
          const document = await parseDocxWithPreviewBudget(
            source,
            { preloadFonts: false },
            starvePreview ? { vmlShape: 0, wpGroup: 0 } : {},
          );
          const authoredCapture = drawingOf(document).rawXml;

          const edited = fromProseDoc(withEditedImage(toProseDoc(document), EDITS[edit]), document);
          const savedDrawing = drawingOf(edited);
          const saved = await savedPackageOf(edited);

          // The drawing survives as authored, and the edit is reported rather
          // than written: a preview folio can still replay is `replayable`, one
          // whose render was edited is `opaque`, and neither is `native`.
          expect(saved.documentXml).toContain(AUTHORED_SIGNATURE[kind]);
          expect(savedDrawing.rawXml).toBe(authoredCapture);
          expect(classifyDrawingSafety(savedDrawing)).not.toBe("native");

          // Nothing folio drew became part of the package.
          expect(saved.entries.some((path) => path.startsWith("word/media/"))).toBe(false);
          expect(saved.relationshipsXml).not.toContain(`Type="${RELATIONSHIP_TYPES.image}"`);
          expect(saved.documentXml).not.toContain("r:embed=");

          // And reopening the save gives back the drawing that went in.
          const reopened = await parseDocx(
            await repackDocx(edited, { updateModifiedDate: false }),
            { preloadFonts: false },
          );
          const reopenedDrawing = drawingOf(reopened);
          expect(reopenedDrawing.rawXml).toBe(authoredCapture);
          expect(reopenedDrawing.image.rId).toBeUndefined();
          expect(expectImageAttrs(imageNodeIn(toProseDoc(reopened)).node).rId).toBeUndefined();
        },
      ),
      propertyConfig(),
    );
  });
});
