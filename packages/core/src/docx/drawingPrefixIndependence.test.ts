/**
 * A drawing is known by its namespace, not by the `wp` prefix.
 *
 * Which of `wp:inline` and `wp:anchor` a `w:drawing` carries is its anchoring,
 * and the WordprocessingDrawing namespace is what says so. The prefix bound to
 * that namespace is the producer's choice: Word writes `wp`, and a package free
 * to bind it to anything else writes the same document. A prefix-matched read
 * found neither element, so such a drawing parsed as no drawing at all — no
 * size, no anchoring, no picture — and a save wrote back whatever the
 * placeholder held.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { createEmptyDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";

const IMAGE_RID = "rIdPicture";

/** A 1×1 PNG, so the drawing has a real picture relationship to resolve. */
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);

/** The same anchored picture, with the WordprocessingDrawing prefix as a parameter. */
const documentXml = (prefix: string): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}" ` +
  `xmlns:a="${A_NAMESPACE}" xmlns:${prefix}="${WP_NAMESPACE}" xmlns:pic="${PIC_NAMESPACE}">` +
  `<w:body><w:p><w:r><w:drawing><${prefix}:anchor distT="3" simplePos="1" relativeHeight="7" ` +
  `behindDoc="0" locked="1" layoutInCell="0" allowOverlap="0" hidden="1">` +
  `<${prefix}:simplePos x="111" y="222"/>` +
  `<${prefix}:positionH relativeFrom="margin"><${prefix}:posOffset>5</${prefix}:posOffset></${prefix}:positionH>` +
  `<${prefix}:positionV relativeFrom="margin"><${prefix}:posOffset>6</${prefix}:posOffset></${prefix}:positionV>` +
  `<${prefix}:extent cx="914400" cy="457200"/>` +
  `<${prefix}:wrapSquare wrapText="left"/>` +
  `<${prefix}:docPr id="1" name="framed" descr="a framed picture"/>` +
  `<a:graphic><a:graphicData uri="${PIC_NAMESPACE}">` +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="${IMAGE_RID}"/></pic:blipFill><pic:spPr/></pic:pic>` +
  `</a:graphicData></a:graphic></${prefix}:anchor></w:drawing></w:r></w:p>` +
  "<w:sectPr/></w:body></w:document>";

const parsedImage = async (prefix: string) => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentXml(prefix));
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
  const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
    preloadFonts: false,
  });
  const paragraph = parsed.package.document.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  const drawing = run?.type === "run" ? run.content.at(0) : undefined;
  return drawing?.type === "drawing" ? drawing.image : undefined;
};

describe("a drawing is resolved by namespace, not by the wp prefix", () => {
  test("a re-prefixed anchor reads the same as Word's", async () => {
    const word = await parsedImage("wp");
    const reprefixed = await parsedImage("dw");

    expect(reprefixed).toEqual(word);
    expect(reprefixed?.size).toEqual({ width: 914_400, height: 457_200 });
    expect(reprefixed?.rId).toBe(IMAGE_RID);
    expect(reprefixed?.docPrName).toBe("framed");
    expect(reprefixed?.alt).toBe("a framed picture");
    expect(reprefixed?.wrap).toEqual({ type: "square", wrapText: "left", distT: 3 });
    expect(reprefixed?.anchor).toEqual({
      useSimplePosition: true,
      locked: true,
      layoutInCell: false,
      allowOverlap: false,
      hidden: true,
      relativeHeight: 7,
      simplePosition: { x: 111, y: 222 },
    });
  });
});
