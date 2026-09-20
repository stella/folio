/**
 * A drawing's graphic-frame locks are written as they were read, or not at all.
 *
 * `serializeGraphicFrameLocks` emitted `noChangeAspect="1"` whenever the model
 * held no lock record, because one `undefined` carried two meanings: the author
 * wrote no `wp:cNvGraphicFramePr`, and folio created this picture itself. At
 * serialization those are indistinguishable, so the default belongs at the
 * insert that creates a picture, and the serializer writes silence back as
 * silence.
 *
 * The property runs the whole lock space — absent, present and empty, and each
 * attribute on its own in both states — through the rebuild path, which is the
 * only path that shows the difference: an untouched document replays its
 * captured drawing bytes and never reaches the serializer.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { Document, DrawingContent, ImageFrameLocks } from "../types/document";
import { GRAPHIC_FRAME_LOCK_KEYS } from "./graphicFrameLocks";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Smallest valid PNG, so the picture resolves to real media. */
const ONE_PIXEL_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const XML_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

const documentXml = (graphicFramePr: string): string => `${XML_DECLARATION}
<w:document ${XML_NAMESPACES}><w:body><w:p><w:r><w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="190500" cy="190500"/>
    <wp:docPr id="7" name="Logo"/>
    ${graphicFramePr}
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic>
    </a:graphicData></a:graphic>
  </wp:inline>
</w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;

const buildDocx = async (graphicFramePr: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
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
      `<Relationship Id="rIdImage" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/></Relationships>`,
    ),
  );
  zip.file("word/media/image1.png", ONE_PIXEL_PNG_BASE64, { base64: true });
  zip.file("word/document.xml", documentXml(graphicFramePr));
  return zip.generateAsync({ type: "arraybuffer" });
};

const firstDrawing = (document: Document): DrawingContent => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected the fixture's first block to be a paragraph");
  }
  const drawing = paragraph.content
    .filter((item) => item.type === "run")
    .flatMap((run) => run.content)
    .find((content) => content.type === "drawing");
  if (drawing === undefined) {
    throw new Error("Expected the fixture paragraph to hold a drawing");
  }
  return drawing;
};

const openDocx = (source: ArrayBuffer): Promise<Document> =>
  parseDocx(source, { detectVariables: false, preloadFonts: false });

/**
 * Every shape the source can carry: no element, the element with no attribute,
 * and each modelled lock on its own in both states.
 */
type FrameState = { xml: string; locks: ImageFrameLocks | undefined };

const FRAME_STATES: FrameState[] = [
  { xml: "", locks: undefined },
  // An authored but empty element says nothing more than no element does, so
  // the model holds the same `undefined` for both.
  {
    xml: "<wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr>",
    locks: undefined,
  },
  ...GRAPHIC_FRAME_LOCK_KEYS.flatMap((key) =>
    [true, false].map((value) => ({
      xml:
        `<wp:cNvGraphicFramePr><a:graphicFrameLocks ${key}="${value ? "1" : "0"}"/>` +
        "</wp:cNvGraphicFramePr>",
      locks: { [key]: value } satisfies ImageFrameLocks,
    })),
  ),
];

describe("a drawing's frame locks survive the rebuild path unchanged", () => {
  test("absent stays absent, and every authored lock comes back", async () => {
    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...FRAME_STATES), async ({ xml, locks }) => {
        const parsed = await openDocx(await buildDocx(xml));
        const drawing = firstDrawing(parsed);
        expect(drawing.image.frameLocks).toEqual(locks);

        // Resizing invalidates the captured drawing XML, so the serializers
        // run: the path every edited document takes.
        drawing.image.size = {
          width: drawing.image.size.width * 2,
          height: drawing.image.size.height * 2,
        };
        const rebuilt = firstDrawing(
          await openDocx(await repackDocx(parsed, { updateModifiedDate: false })),
        );

        expect(rebuilt.image.frameLocks).toEqual(locks);
      }),
      propertyConfig({ numRuns: FRAME_STATES.length * 2 }),
    );
  });
});
