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

import type { DrawingContent, Paragraph, Run, Table } from "../types/document";
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

/** Pull the first drawing content out of a paragraph's first run. */
function firstDrawing(block: Paragraph | Table | undefined): DrawingContent | undefined {
  if (block?.type !== "paragraph") {
    return undefined;
  }
  const run = block.content.find((c): c is Run => c.type === "run");
  return run?.content.find((c): c is DrawingContent => c.type === "drawing");
}

/** First drawing carried by a run nested inside the paragraph's first hyperlink. */
function firstDrawingInHyperlink(block: Paragraph | Table | undefined): DrawingContent | undefined {
  if (block?.type !== "paragraph") {
    return undefined;
  }
  const hyperlink = block.content.find((c) => c.type === "hyperlink");
  if (hyperlink?.type !== "hyperlink") {
    return undefined;
  }
  for (const child of hyperlink.children) {
    if (child.type !== "run") {
      continue;
    }
    const drawing = child.content.find((c): c is DrawingContent => c.type === "drawing");
    if (drawing) {
      return drawing;
    }
  }
  return undefined;
}

/** First drawing carried by a run nested inside the paragraph's first inline SDT. */
function firstDrawingInInlineSdt(block: Paragraph | Table | undefined): DrawingContent | undefined {
  if (block?.type !== "paragraph") {
    return undefined;
  }
  const sdt = block.content.find((c) => c.type === "inlineSdt");
  if (sdt?.type !== "inlineSdt") {
    return undefined;
  }
  for (const child of sdt.content) {
    if (child.type !== "run") {
      continue;
    }
    const drawing = child.content.find((c): c is DrawingContent => c.type === "drawing");
    if (drawing) {
      return drawing;
    }
  }
  return undefined;
}

/**
 * Build a DOCX whose document root declares only the `w` namespace plus a single
 * image relationship (`rIdImg`). The caller supplies the `<w:body>` inner XML so
 * a test can scope non-canonical VML / relationship prefixes on whichever
 * wrapper element (`w:hyperlink`, inline `w:sdt`, …) it is exercising.
 */
async function bodyScopedPictDocx(bodyInnerXml: string): Promise<ArrayBuffer> {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyInnerXml}
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
  return zip.generateAsync({ type: "arraybuffer" });
}

const PICT_WITH_IMAGE = `<w:pict><v:shape id="Picture 1" type="#_x0000_t75" style="width:2in;height:1in"><v:imagedata r:id="rIdImg" o:title="logo"/></v:shape></w:pict>`;
const OBJECT_WITH_PREVIEW = `<w:object><v:shape id="Object 1" type="#_x0000_t75" style="width:20pt;height:18pt"><v:imagedata r:id="rIdImg" o:title="preview"/></v:shape><w:control r:id="rIdControl" w:name="Control 1" w:shapeid="Object 1"/></w:object>`;

// A VML shape/imagedata whose prefixes (`v2`, `r2`) are declared on an ancestor
// wrapper rather than the document root; used to prove the in-scope xmlns set
// accumulates through that wrapper before the captured pict is emitted.
const PICT_WITH_ALT_PREFIXES = `<w:pict><v2:shape style="width:2in;height:1in"><v2:imagedata r2:id="rIdImg"/></v2:shape></w:pict>`;

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

  test("parses an embedded object's VML preview with its authored dimensions", async () => {
    const doc = await parseDocx(await pictDocx({ runXml: OBJECT_WITH_PREVIEW }), {
      preloadFonts: false,
    });

    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.size).toEqual({ width: 254_000, height: 228_600 });
    expect(drawing?.image.wrap.type).toBe("inline");
    expect(drawing?.image.src).toStartWith("data:image/png");
    expect(drawing?.image.title).toBe("preview");
    expect(drawing?.rawXml).toContain("<w:object");
    expect(drawing?.rawXml).toContain("<w:control");
  });

  test("preserves an embedded object wrapper when saving its preview", async () => {
    const original = await pictDocx({ runXml: OBJECT_WITH_PREVIEW });
    const doc = await parseDocx(original, { preloadFonts: false });
    const out = await repackDocx(doc, { updateModifiedDate: false });

    expect((await validateDocx(out)).valid).toBe(true);
    const zip = await JSZip.loadAsync(out);
    const docXml = await zip.file("word/document.xml")!.async("text");
    expect(docXml).toContain("<w:object");
    expect(docXml).toContain("<w:control");
    expect(docXml).toContain('r:id="rIdImg"');
    expect(docXml).not.toContain("<w:drawing");
  });

  test("rejects an embedded-object preview with unsafe authored dimensions", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: OBJECT_WITH_PREVIEW.replace("width:20pt", "width:30000px"),
      }),
      { preloadFonts: false },
    );

    expect(firstDrawing(doc.package.document.content.at(0))).toBeUndefined();
  });

  test("renders a positioned solid rectangle without embedded image data", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:rect style="position:absolute;margin-left:12pt;margin-top:6pt;width:1in;height:.5in;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" fillcolor="black" stroked="f"/></w:pict>`,
      }),
      { preloadFonts: false },
    );

    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.size).toEqual({ width: 914_400, height: 457_200 });
    expect(drawing?.image.wrap.type).toBe("inFront");
    expect(drawing?.image.position).toEqual({
      horizontal: { relativeTo: "page", posOffset: 152_400 },
      vertical: { relativeTo: "page", posOffset: 76_200 },
    });
    expect(drawing?.image.src).toStartWith("data:image/svg+xml");
    expect(decodeURIComponent(drawing?.image.src?.split(",").at(1) ?? "")).toContain(
      'fill="black"',
    );
    expect(drawing?.rawXml).toContain("<w:pict");
  });

  test("renders bounded rectangle and line children from a positioned VML group", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:group style="position:absolute;margin-left:10pt;margin-top:20pt;width:100pt;height:50pt;z-index:-1;mso-position-horizontal-relative:page" coordorigin="100,200" coordsize="1000,500"><v:line from="100,200" to="1100,700"/><v:rect style="left:200;top:250;width:400;height:100" fillcolor="#112233" stroked="f"/></v:group></w:pict>`,
      }),
      { preloadFonts: false },
    );

    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.size).toEqual({ width: 1_270_000, height: 635_000 });
    expect(drawing?.image.wrap.type).toBe("behind");
    expect(drawing?.image.position?.horizontal).toEqual({
      relativeTo: "page",
      posOffset: 127_000,
    });
    expect(drawing?.image.position?.vertical).toEqual({
      relativeTo: "paragraph",
      posOffset: 254_000,
    });
    const svg = decodeURIComponent(drawing?.image.src?.split(",").at(1) ?? "");
    expect(svg).toContain('viewBox="100 200 1000 500"');
    expect(svg).toContain("<line");
    expect(svg).toContain("<rect");
  });

  test("skips solid-shape previews with unsafe dimensions", async () => {
    const doc = await parseDocx(
      await pictDocx({
        runXml: `<w:pict><v:rect style="width:999999px;height:1in" fillcolor="black"/></w:pict>`,
      }),
      { preloadFonts: false },
    );

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

  test("uses the Choice image when the Fallback pict's relationship cannot be resolved", async () => {
    // The Fallback VML references a missing relationship, so it cannot resolve
    // to a media part; the resolvable Choice DrawingML image must win.
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
    // Only the Choice image resolves; the Fallback's rIdMissing is absent.
    zip.file(
      "word/_rels/document.xml.rels",
      `${XML}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChoice" Type="${RELATIONSHIP_TYPES.image}" Target="media/choice.png"/>
</Relationships>`,
    );
    zip.file(
      "word/document.xml",
      `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="choice"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="choice"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdChoice"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>
        <mc:Fallback><w:pict><v:shape style="width:2in;height:1in"><v:imagedata r:id="rIdMissing"/></v:shape></w:pict></mc:Fallback>
      </mc:AlternateContent>
    </w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
    );
    zip.file("word/media/choice.png", ONE_PIXEL_PNG_BASE64, { base64: true });
    zip.file(
      "word/styles.xml",
      `${XML}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`,
    );
    const original = await zip.generateAsync({ type: "arraybuffer" });

    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawing(doc.package.document.content.at(0));
    // The resolvable Choice image is used, not the unresolvable Fallback pict.
    expect(drawing?.image.rId).toBe("rIdChoice");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
  });

  test("captures xmlns declarations scoped on a table cell containing a w:pict", async () => {
    // The non-canonical VML / relationship prefixes are bound on the <w:tc>, so
    // the in-scope set must accumulate through tbl -> tr -> tc into the cell.
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
    // Root declares only w; v2/r2 are scoped on the <w:tc>.
    zip.file(
      "word/document.xml",
      `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr><w:tc xmlns:v2="urn:schemas-microsoft-com:vml" xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
        <w:p><w:r><w:pict><v2:shape style="width:2in;height:1in"><v2:imagedata r2:id="rIdImg"/></v2:shape></w:pict></w:r></w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:p/>
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
    const table = doc.package.document.content.at(0);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a table");
    }
    const drawing = firstDrawing(table.rows.at(0)?.cells.at(0)?.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const outZip = await JSZip.loadAsync(out);
    const docXml = await outZip.file("word/document.xml")!.async("text");
    // The cell-scoped bindings are carried onto the captured pict.
    expect(docXml).toContain('xmlns:v2="urn:schemas-microsoft-com:vml"');
    expect(docXml).toContain("<v2:shape");
    expect(docXml).toContain('r2:id="rIdImg"');
  });

  test("captures xmlns declarations scoped on a w:hyperlink wrapping a w:pict", async () => {
    // The non-canonical VML / relationship prefixes are bound on the
    // <w:hyperlink>, so the in-scope set must accumulate through the hyperlink
    // wrapper into its child run before the captured pict is emitted.
    const original = await bodyScopedPictDocx(
      `<w:p><w:hyperlink w:anchor="section1" xmlns:v2="urn:schemas-microsoft-com:vml" xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r>${PICT_WITH_ALT_PREFIXES}</w:r></w:hyperlink></w:p>`,
    );

    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawingInHyperlink(doc.package.document.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
    expect(drawing?.image.size.width).toBe(1_828_800);

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const outZip = await JSZip.loadAsync(out);
    const docXml = await outZip.file("word/document.xml")!.async("text");
    // The hyperlink-scoped bindings are carried onto the captured pict.
    expect(docXml).toContain('xmlns:v2="urn:schemas-microsoft-com:vml"');
    expect(docXml).toContain("<v2:shape");
    expect(docXml).toContain('r2:id="rIdImg"');

    // Re-parsing the saved output still resolves the picture (lossless).
    const reparsed = await parseDocx(out, { preloadFonts: false });
    const reDrawing = firstDrawingInHyperlink(reparsed.package.document.content.at(0));
    expect(reDrawing?.image.rId).toBe("rIdImg");
    expect(reDrawing?.image.src?.startsWith("data:image/png")).toBe(true);
  });

  test("captures xmlns declarations scoped on an inline w:sdt wrapping a w:pict", async () => {
    // The non-canonical prefixes are bound on the <w:sdt> element itself (not on
    // its <w:sdtContent>), so the merge must fold the sdt wrapper's own xmlns in
    // before recursing into the content control.
    const original = await bodyScopedPictDocx(
      `<w:p><w:sdt xmlns:v2="urn:schemas-microsoft-com:vml" xmlns:r2="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:sdtPr><w:tag w:val="logo"/></w:sdtPr><w:sdtContent><w:r>${PICT_WITH_ALT_PREFIXES}</w:r></w:sdtContent></w:sdt></w:p>`,
    );

    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawingInInlineSdt(doc.package.document.content.at(0));
    expect(drawing?.image.rId).toBe("rIdImg");
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
    expect(drawing?.image.size.width).toBe(1_828_800);

    const out = await repackDocx(doc, { updateModifiedDate: false });
    expect((await validateDocx(out)).valid).toBe(true);

    const outZip = await JSZip.loadAsync(out);
    const docXml = await outZip.file("word/document.xml")!.async("text");
    // The sdt-scoped bindings are carried onto the captured pict.
    expect(docXml).toContain('xmlns:v2="urn:schemas-microsoft-com:vml"');
    expect(docXml).toContain("<v2:shape");
    expect(docXml).toContain('r2:id="rIdImg"');

    // Re-parsing the saved output still resolves the picture (lossless).
    const reparsed = await parseDocx(out, { preloadFonts: false });
    const reDrawing = firstDrawingInInlineSdt(reparsed.package.document.content.at(0));
    expect(reDrawing?.image.rId).toBe("rIdImg");
    expect(reDrawing?.image.src?.startsWith("data:image/png")).toBe(true);
  });
});

describe("VML style attribute hardening", () => {
  test("ignores prototype-polluting style keys without breaking width parse", async () => {
    const marker = "__folio_vml_style_pollution__";
    expect(Object.hasOwn(Object.prototype, marker)).toBe(false);

    const original = await pictDocx({
      runXml: `<w:pict><v:shape style="width:120pt;height:40pt;__proto__:${marker};constructor:x;prototype:y"><v:imagedata r:id="rIdImg"/></v:shape></w:pict>`,
    });
    const doc = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawing(doc.package.document.content.at(0));
    expect(drawing?.image.src?.startsWith("data:image/png")).toBe(true);
    expect(drawing?.image.size.width).toBeGreaterThan(0);

    expect(Object.hasOwn(Object.prototype, marker)).toBe(false);
    expect(({} as Record<string, unknown>)[marker]).toBeUndefined();
  });
});
