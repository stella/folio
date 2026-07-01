/**
 * Legacy VML `w:pict` inline images.
 *
 * Older Word files embed inline pictures (e.g. header logos) as VML instead of
 * DrawingML. The run parser used to drop `w:pict`, so these never rendered.
 * These tests cover: resolving `<v:imagedata r:id>` to the same drawing/image
 * node a DrawingML image produces (with style-derived dimensions and resolved
 * media), byte-preserving the original VML on save, and skipping degenerate
 * shapes (no imagedata, unresolved relationship) without throwing.
 *
 * Ported alongside eigenpal/docx-editor `vmlImageParser.ts`.
 */
import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { DrawingContent, Paragraph, Run } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx, validateDocx } from "./rezip";

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ONE_PIXEL_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

type PictDocxOptions = {
  /** The `<w:r>` inner XML placed in the body paragraph. */
  runXml: string;
  /** When false, omit the image relationship so the r:id does not resolve. */
  imageRel?: boolean;
  /** When false, omit the media part. */
  media?: boolean;
};

async function pictDocx(options: PictDocxOptions): Promise<ArrayBuffer> {
  const imageRel = options.imageRel ?? true;
  const withMedia = options.media ?? true;
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${imageRel ? `  <Relationship Id="rIdImg" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>` : ""}
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r>${options.runXml}</w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  if (withMedia) {
    zip.file("word/media/image1.png", ONE_PIXEL_PNG_BASE64, { base64: true });
  }
  zip.file(
    "word/styles.xml",
    `${XML}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
}

/** Pull the first drawing content out of the first body paragraph's first run. */
function firstDrawing(paragraph: Paragraph | undefined): DrawingContent | undefined {
  if (paragraph?.type !== "paragraph") {
    return undefined;
  }
  const run = paragraph.content.find((c): c is Run => c.type === "run");
  return run?.content.find((c): c is DrawingContent => c.type === "drawing");
}

const PICT_WITH_IMAGE = `<w:pict><v:shape id="Picture 1" type="#_x0000_t75" style="width:2in;height:1in"><v:imagedata r:id="rIdImg" o:title="logo"/></v:shape></w:pict>`;

describe("VML w:pict inline images", () => {
  test("parses a w:pict into a drawing/image node with style dimensions + resolved media", async () => {
    const doc = await parseDocx(await pictDocx({ runXml: PICT_WITH_IMAGE }), {
      preloadFonts: false,
    });

    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing).toBeDefined();
    expect(drawing?.image.type).toBe("image");
    expect(drawing?.image.rId).toBe("rIdImg");

    // 2in = 192px = 1,828,800 EMU; 1in = 96px = 914,400 EMU (1px = 9525 EMU).
    expect(drawing?.image.size.width).toBe(1_828_800);
    expect(drawing?.image.size.height).toBe(914_400);
    expect(drawing?.image.wrap.type).toBe("inline");

    // The <v:imagedata r:id> resolved to the media part's bytes.
    expect(drawing?.image.src).toBeDefined();
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
    expect(drawing?.image.mimeType).toBe("image/png");
    expect(drawing?.image.filename).toBe("image1.png");
    expect(drawing?.image.title).toBe("logo");
  });

  test("preserves the original VML verbatim on save (no DrawingML conversion)", async () => {
    const original = await pictDocx({ runXml: PICT_WITH_IMAGE });
    const doc = await parseDocx(original, { preloadFonts: false });

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const zip = await JSZip.loadAsync(out);
    const docXml = await zip.file("word/document.xml")!.async("text");
    // The VML round-trips as VML, not a synthesized <w:drawing>.
    expect(docXml).toContain("<w:pict");
    expect(docXml).toContain("<v:imagedata");
    expect(docXml).toContain('r:id="rIdImg"');
    expect(docXml).not.toContain("<w:drawing");
    expect(docXml).not.toContain("wp:inline");
    // The referenced media part survives the round-trip.
    expect(zip.file("word/media/image1.png")).not.toBeNull();
  });

  test("skips a w:pict with no imagedata without throwing", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:rect style="width:1in;height:1in"/></w:pict>`,
      }),
      { preloadFonts: false },
    );

    // No image node is produced; the run simply carries no drawing content.
    expect(firstDrawing(doc.package.document.content.at(0))).toBeUndefined();
  });

  test("keeps a w:pict whose relationship does not resolve (no src) and preserves it on save", async () => {
    const doc = await parseDocx(await pictDocx({ runXml: PICT_WITH_IMAGE, imageRel: false }), {
      preloadFonts: false,
    });

    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing).toBeDefined();
    expect(drawing?.image.rId).toBe("rIdImg");
    // Unresolved relationship: no binary, but the reference is not dropped.
    expect(drawing?.image.src).toBeUndefined();

    const out = await repackDocx(doc, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(out);
    const docXml = await zip.file("word/document.xml")!.async("text");
    expect(docXml).toContain("<w:pict");
    expect(docXml).toContain('r:id="rIdImg"');
  });
});
