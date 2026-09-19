import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";

import { expectImageAttrs, mergeImageAttrs } from "../prosemirror/attrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, DrawingContent } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

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

const picturePartsXml = (graphicFramePr: string): string => `<w:p><w:r><w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="190500" cy="190500"/>
    <wp:effectExtent l="12700" t="19050" r="6350" b="25400"/>
    <wp:docPr id="7" name="Logo" descr="Mark" title="Company logo"/>
    ${graphicFramePr}
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic>
    </a:graphicData></a:graphic>
  </wp:inline>
</w:drawing></w:r></w:p>`;

/** A body paragraph holding one inline picture, with the given frame properties. */
const createPictureDocx = async (graphicFramePr: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
</Relationships>`,
  );
  zip.file("word/media/image1.png", ONE_PIXEL_PNG_BASE64, { base64: true });
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document ${XML_NAMESPACES}>
  <w:body>
    ${picturePartsXml(graphicFramePr)}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
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

/** Resizing invalidates the raw-XML fingerprint, forcing DrawingML regeneration. */
const resize = (drawing: DrawingContent): void => {
  drawing.image.size = {
    width: drawing.image.size.width * 2,
    height: drawing.image.size.height * 2,
  };
};

const AUTHORED_FRAME_PR =
  '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="0" noMove="1"/></wp:cNvGraphicFramePr>';

const AUTHORED_LOCKS = { noChangeAspect: false, noMove: true };

/** The fixture's `wp:effectExtent`, in EMU. */
const AUTHORED_PADDING = { left: 12_700, top: 19_050, right: 6350, bottom: 25_400 };

/** Everything a regenerated picture must still carry, other than its size. */
const expectAuthoredFidelity = (drawing: DrawingContent): void => {
  expect(drawing.image.frameLocks).toEqual(AUTHORED_LOCKS);
  expect(drawing.image.padding).toEqual(AUTHORED_PADDING);
  expect(drawing.image.docPrName).toBe("Logo");
  expect(drawing.image.alt).toBe("Mark");
  expect(drawing.image.title).toBe("Company logo");
};

describe("picture regeneration fidelity", () => {
  test("keeps frame locks, effect extent and docPr across a model resize", async () => {
    const parsed = await openDocx(await createPictureDocx(AUTHORED_FRAME_PR));
    const drawing = firstDrawing(parsed);

    // Without raw XML there is nothing for the resize to invalidate, so the
    // regeneration this test covers would never run.
    expect(drawing.rawXml).toBeDefined();
    expect(drawing.image.size).toEqual({ width: 190_500, height: 190_500 });
    expectAuthoredFidelity(drawing);

    resize(drawing);
    const reopened = firstDrawing(
      await openDocx(await repackDocx(parsed, { updateModifiedDate: false })),
    );

    expect(reopened.image.size).toEqual({ width: 381_000, height: 381_000 });
    expectAuthoredFidelity(reopened);
  });

  test("keeps frame locks, effect extent and docPr across an in-editor resize", async () => {
    const parsed = await openDocx(await createPictureDocx(AUTHORED_FRAME_PR));

    const pmDocument = toProseDoc(parsed);
    let imagePosition: number | undefined;
    pmDocument.descendants((node, position) => {
      if (node.type.name === "image") {
        imagePosition = position;
      }
    });
    if (imagePosition === undefined) {
      throw new Error("Expected the fixture to produce an editor image node");
    }

    const state = EditorState.create({ doc: pmDocument });
    const imageNode = state.doc.nodeAt(imagePosition);
    if (!imageNode) {
      throw new Error("Expected an image node at the resolved position");
    }
    const { width, height } = expectImageAttrs(imageNode);
    if (width === undefined || height === undefined) {
      throw new Error("Expected the fixture image to carry a size");
    }
    const resized = state.apply(
      state.tr.setNodeMarkup(
        imagePosition,
        null,
        mergeImageAttrs(imageNode, { width: width * 2, height: height * 2 }),
      ),
    ).doc;

    const saved = await repackDocx(fromProseDoc(resized, parsed), { updateModifiedDate: false });
    const reopened = firstDrawing(await openDocx(saved));

    expect(reopened.image.size).toEqual({ width: 381_000, height: 381_000 });
    expectAuthoredFidelity(reopened);
  });

  test("writes no frame properties when the picture authored none", async () => {
    const parsed = await openDocx(await createPictureDocx(""));
    const drawing = firstDrawing(parsed);

    expect(drawing.image.frameLocks).toBeUndefined();

    resize(drawing);
    const reopened = firstDrawing(
      await openDocx(await repackDocx(parsed, { updateModifiedDate: false })),
    );

    expect(reopened.image.size).toEqual({ width: 381_000, height: 381_000 });
    // A lock the author never wrote is not invented by the rebuild.
    expect(reopened.image.frameLocks).toBeUndefined();
  });
});
