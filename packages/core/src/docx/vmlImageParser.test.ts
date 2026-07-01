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

  test("parses shape dimensions with case-insensitive units", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:shape style="width:2IN;height:1Pt"><v:imagedata r:id="rIdImg"/></v:shape></w:pict>`,
      }),
      { preloadFonts: false },
    );

    const drawing = firstDrawing(doc.package.document.content.at(0));
    // 2IN = 192px = 1,828,800 EMU; 1Pt = 96/72 px = 12,700 EMU.
    expect(drawing?.image.size.width).toBe(1_828_800);
    expect(drawing?.image.size.height).toBe(12_700);
  });

  test("rejects a malformed length rather than mis-parsing it", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:shape style="width:1.2.3in;height:1in"><v:imagedata r:id="rIdImg"/></v:shape></w:pict>`,
      }),
      { preloadFonts: false },
    );

    const drawing = firstDrawing(doc.package.document.content.at(0));
    // `1.2.3in` is not a valid length, so the width is dropped (0), not truncated
    // to `1.2`; the valid height still parses.
    expect(drawing?.image.size.width).toBe(0);
    expect(drawing?.image.size.height).toBe(914_400);
  });

  test("resolves the relationship id from o:relid when r:id is absent", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:shape style="width:1in;height:1in"><v:imagedata o:relid="rIdImg"/></v:shape></w:pict>`,
      }),
      { preloadFonts: false },
    );

    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
  });

  test("round-trips a w:pict that binds VML/relationship namespaces to non-canonical prefixes", async () => {
    // A DOCX whose root binds VML (`v2`), office (`o2`), and relationship (`r2`)
    // namespaces to non-canonical prefixes the serializer's root does not
    // declare. The captured VML must carry those declarations so it resolves.
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
  <Relationship Id="rIdImg" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
</Relationships>`,
    );
    zip.file(
      "word/document.xml",
      `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v2="urn:schemas-microsoft-com:vml" xmlns:o2="urn:schemas-microsoft-com:office:office" xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:pict><v2:shape style="width:2in;height:1in"><v2:imagedata r2:id="rIdImg" o2:title="logo"/></v2:shape></w:pict></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
    );
    zip.file("word/media/image1.png", ONE_PIXEL_PNG_BASE64, { base64: true });
    zip.file(
      "word/styles.xml",
      `${XML}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`,
    );
    const original = await zip.generateAsync({ type: "arraybuffer" });

    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
    expect(drawing?.image.size.width).toBe(1_828_800);

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const outZip = await JSZip.loadAsync(out);
    const docXml = await outZip.file("word/document.xml")!.async("text");
    // The non-canonical prefix bindings travel with the captured VML.
    expect(docXml).toContain('xmlns:v2="urn:schemas-microsoft-com:vml"');
    expect(docXml).toContain("<v2:shape");
    expect(docXml).toContain('r2:id="rIdImg"');

    // Re-parsing the saved output still resolves the picture (lossless).
    const reparsed = await parseDocx(out, { preloadFonts: false });
    const reDrawing = firstDrawing(reparsed.package.document.content.at(0));
    expect(reDrawing?.image.rId).toBe("rIdImg");
    expect(reDrawing?.image.src?.startsWith("data:image/png")).toBe(true);
  });

  test("routes a w:pict inside mc:AlternateContent/mc:Fallback and prefers the fallback image", async () => {
    // A common compatibility form: a modern DrawingML Choice plus a VML `w:pict`
    // Fallback carrying the picture. folio renders (and keeps) the Fallback image.
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
  <Relationship Id="rIdChoice" Type="${RELATIONSHIP_TYPES.image}" Target="media/choice.png"/>
  <Relationship Id="rIdFallback" Type="${RELATIONSHIP_TYPES.image}" Target="media/fallback.png"/>
</Relationships>`,
    );
    zip.file(
      "word/document.xml",
      `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdChoice"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>
        <mc:Fallback><w:pict><v:shape style="width:2in;height:1in"><v:imagedata r:id="rIdFallback" o:title="fallback"/></v:shape></w:pict></mc:Fallback>
      </mc:AlternateContent>
    </w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
    );
    zip.file("word/media/choice.png", ONE_PIXEL_PNG_BASE64, { base64: true });
    zip.file("word/media/fallback.png", ONE_PIXEL_PNG_BASE64, { base64: true });
    zip.file(
      "word/styles.xml",
      `${XML}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`,
    );
    const original = await zip.generateAsync({ type: "arraybuffer" });

    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawing(doc.package.document.content.at(0));
    // The Fallback's VML picture is preferred over the Choice's DrawingML image.
    expect(drawing?.image.rId).toBe("rIdFallback");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
    expect(drawing?.image.size.width).toBe(1_828_800);

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const outZip = await JSZip.loadAsync(out);
    const docXml = await outZip.file("word/document.xml")!.async("text");
    expect(docXml).toContain("<w:pict");
    expect(docXml).toContain('r:id="rIdFallback"');
    // The whole mc:AlternateContent is preserved on save (the Choice is not lost).
    expect(docXml).toContain('r:embed="rIdChoice"');
    expect(outZip.file("word/media/fallback.png")).not.toBeNull();
  });

  test("captures xmlns declarations scoped on an ancestor paragraph (not just the root)", async () => {
    // The non-canonical VML / relationship prefixes are bound on the <w:p>, not
    // the document root. The in-scope set must accumulate down the ancestor
    // chain so the captured pict still carries them.
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
  <Relationship Id="rIdImg" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
</Relationships>`,
    );
    // Root declares only w; v2/r2 are scoped on the paragraph.
    zip.file(
      "word/document.xml",
      `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p xmlns:v2="urn:schemas-microsoft-com:vml" xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:pict><v2:shape style="width:2in;height:1in"><v2:imagedata r2:id="rIdImg"/></v2:shape></w:pict></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
    );
    zip.file("word/media/image1.png", ONE_PIXEL_PNG_BASE64, { base64: true });
    zip.file(
      "word/styles.xml",
      `${XML}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`,
    );
    const original = await zip.generateAsync({ type: "arraybuffer" });

    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const outZip = await JSZip.loadAsync(out);
    const docXml = await outZip.file("word/document.xml")!.async("text");
    // The paragraph-scoped bindings are carried onto the captured pict.
    expect(docXml).toContain('xmlns:v2="urn:schemas-microsoft-com:vml"');
    expect(docXml).toContain("<v2:shape");
    expect(docXml).toContain('r2:id="rIdImg"');

    const reparsed = await parseDocx(out, { preloadFonts: false });
    const reDrawing = firstDrawing(reparsed.package.document.content.at(0));
    expect(reDrawing?.image.rId).toBe("rIdImg");
    expect(reDrawing?.image.src?.startsWith("data:image/png")).toBe(true);
  });
});
