import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import {
  type TrackedSectionEndpointRemoval,
  withTrackedSectionEndpointRemoval,
} from "../internal/sectionEndpointResolution";
import type { Document, Image } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createDocx, createEmptyDocx, DocxPackageFidelityError, repackDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const ONE_PIXEL_PNG_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

const URI_ENCODED_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>';

const repackAfterTrackedSectionEndpointRemoval = ({
  document,
  resolution,
}: {
  document: Document;
  resolution: TrackedSectionEndpointRemoval;
}): Promise<ArrayBuffer> =>
  withTrackedSectionEndpointRemoval({
    document,
    resolution,
    repack: () => repackDocx(document, { updateModifiedDate: false }),
  });

async function createHeaderFixture(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/header1.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
}

function truncateAfterEndOfCentralDirectoryCounts(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  for (let offset = bytes.byteLength - 22; offset >= 0; offset -= 1) {
    if (
      bytes[offset] === 0x50 &&
      bytes[offset + 1] === 0x4b &&
      bytes[offset + 2] === 0x05 &&
      bytes[offset + 3] === 0x06
    ) {
      return bytes.slice(0, offset + 12).buffer;
    }
  }

  throw new Error("Could not find ZIP end-of-central-directory record");
}

async function createAlternateContentImageFixture(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
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
  <Relationship Id="rId5" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.emf"/>
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p w14:paraId="IMG00001">
      <w:r>
        <mc:AlternateContent>
          <mc:Choice Requires="wps">
            <w:drawing>
              <wp:inline>
                <wp:extent cx="9525" cy="9525"/>
                <wp:docPr id="7" name="Unsupported vector image"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                      <pic:nvPicPr><pic:cNvPr id="7" name="image1.emf"/><pic:cNvPicPr/></pic:nvPicPr>
                      <pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                    </pic:pic>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </mc:Choice>
        </mc:AlternateContent>
      </w:r>
    </w:p>
    <w:p w14:paraId="TXT00001"><w:r><w:t>Editable text</w:t></w:r></w:p>
  </w:body>
</w:document>`,
  );
  zip.file("word/media/image1.emf", new Uint8Array([1, 2, 3, 4]));
  return zip.generateAsync({ type: "arraybuffer" });
}

async function createHyperlinkedImageFixture(): Promise<ArrayBuffer> {
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
  <Relationship Id="rId5" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
  <Relationship Id="rId6" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="https://example.test/guide" TargetMode="External"/>
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="9525" cy="9525"/>
            <wp:docPr id="7" name="image1.png"><a:hlinkClick r:id="rId6"/></wp:docPr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="7" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`,
  );
  zip.file("word/media/image1.png", new Uint8Array([1, 2, 3, 4]));
  return zip.generateAsync({ type: "arraybuffer" });
}

function getFirstImage(document: Document): Image {
  const paragraph = document.package.document.content.at(0);
  if (!paragraph || paragraph.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }
  const run = paragraph.content.at(0);
  if (!run || run.type !== "run") {
    throw new Error("expected a run");
  }
  const drawing = run.content.at(0);
  if (!drawing || drawing.type !== "drawing") {
    throw new Error("expected a drawing");
  }
  return drawing.image;
}

async function createMultiSectionFirstHeaderImageFixture(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
  <Relationship Id="rId11" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>
  <Relationship Id="rId12" Type="${RELATIONSHIP_TYPES.header}" Target="header2.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/header1.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="2AA00001">
      <w:r><w:t>First section text</w:t></w:r>
      <w:pPr>
        <w:sectPr>
          <w:headerReference w:type="first" r:id="rId11"/>
          <w:titlePg/>
        </w:sectPr>
      </w:pPr>
    </w:p>
    <w:p w14:paraId="3BB00001"><w:r><w:t>Second section text</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId12"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="image1.png"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
</w:hdr>`,
  );
  zip.file(
    "word/header2.xml",
    `${XML_DECLARATION}
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Default header</w:t></w:r></w:p></w:hdr>`,
  );
  zip.file("word/media/image1.png", new Uint8Array([1, 2, 3, 4]));
  return zip.generateAsync({ type: "arraybuffer" });
}

describe("repackDocx", () => {
  test("ignores foreign-namespace section-like elements in fidelity checks", async () => {
    const source = await createEmptyDocx();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const sourceZip = await JSZip.loadAsync(source);
    const documentFile = sourceZip.file("word/document.xml");
    if (!documentFile) {
      throw new Error("expected word/document.xml");
    }
    const documentXml = await documentFile.async("text");
    sourceZip.file(
      "word/document.xml",
      documentXml.replace(
        "</w:body>",
        '<foreign:sectPr xmlns:foreign="urn:example:foreign"/></w:body>',
      ),
    );
    parsed.originalBuffer = await sourceZip.generateAsync({ type: "arraybuffer" });

    await expect(repackDocx(parsed, { updateModifiedDate: false })).resolves.toBeInstanceOf(
      ArrayBuffer,
    );
  });

  test("counts Strict WordprocessingML section elements in fidelity checks", async () => {
    const source = await createEmptyDocx();
    const parsed = await parseDocx(source, { preloadFonts: false });
    const sourceZip = await JSZip.loadAsync(source);
    const documentFile = sourceZip.file("word/document.xml");
    if (!documentFile) {
      throw new Error("expected word/document.xml");
    }
    const documentXml = await documentFile.async("text");
    sourceZip.file(
      "word/document.xml",
      documentXml.replace(
        "</w:body>",
        '<strict:sectPr xmlns:strict="http://purl.oclc.org/ooxml/wordprocessingml/main"/></w:body>',
      ),
    );
    parsed.originalBuffer = await sourceZip.generateAsync({ type: "arraybuffer" });

    await expect(repackDocx(parsed, { updateModifiedDate: false })).rejects.toBeInstanceOf(
      DocxPackageFidelityError,
    );
  });

  test("preserves an explicit false final-section bidi setting across save and reopen", async () => {
    const sourceZip = await JSZip.loadAsync(await createEmptyDocx());
    const documentFile = sourceZip.file("word/document.xml");
    if (!documentFile) {
      throw new Error("expected word/document.xml");
    }
    const documentXml = await documentFile.async("text");
    const explicitLtrXml = documentXml.replace("</w:sectPr>", '<w:bidi w:val="0"/></w:sectPr>');
    expect(explicitLtrXml).not.toBe(documentXml);
    sourceZip.file("word/document.xml", explicitLtrXml);

    const parsed = await parseDocx(await sourceZip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    expect(parsed.package.document.finalSectionProperties?.bidi).toBe(false);

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const reopened = await parseDocx(saved, { preloadFonts: false });

    expect(reopened.package.document.finalSectionProperties?.bidi).toBe(false);
    expect(reopened.package.document.sections?.at(-1)?.properties.bidi).toBe(false);
  });

  test("preserves clickable image hyperlinks across full save and reopen", async () => {
    const parsed = await parseDocx(await createHyperlinkedImageFixture(), {
      preloadFonts: false,
    });
    expect(getFirstImage(parsed)).toMatchObject({
      hlinkHref: "https://example.test/guide",
      hlinkRId: "rId6",
    });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const savedZip = await JSZip.loadAsync(saved);
    expect(await savedZip.file("word/document.xml")?.async("text")).toContain(
      '<a:hlinkClick r:id="rId6"/>',
    );

    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(getFirstImage(reopened)).toMatchObject({
      hlinkHref: "https://example.test/guide",
      hlinkRId: "rId6",
    });
    expect(reopened.package.document.content).toEqual(parsed.package.document.content);
  });

  test("writes URI-encoded SVG data URLs as image parts", async () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    {
                      type: "drawing",
                      image: {
                        type: "image",
                        src: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(URI_ENCODED_SVG)}`,
                        filename: "shape.svg",
                        alt: "Shape",
                        size: { width: 95_250, height: 95_250 },
                        wrap: { type: "inline" },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const result = await createDocx(document);
    const zip = await JSZip.loadAsync(result);

    expect(await zip.file("word/media/image1.svg")?.async("text")).toBe(URI_ENCODED_SVG);
    expect(await zip.file("word/document.xml")?.async("text")).toContain('r:embed="rId2"');
    expect(await zip.file("[Content_Types].xml")?.async("text")).toContain(
      'Extension="svg" ContentType="image/svg+xml"',
    );
  });

  test("reuses the cached source package while preserving later model edits", async () => {
    const sourceZip = await JSZip.loadAsync(await createHeaderFixture());
    sourceZip.file("word/vbaProject.bin", new Uint8Array([1, 2, 3]));
    const sourceBuffer = await sourceZip.generateAsync({ type: "arraybuffer" });
    const document = await parseDocx(sourceBuffer, { preloadFonts: false });

    const firstSave = await createDocx(document);
    expect(
      await (await JSZip.loadAsync(firstSave)).file("word/vbaProject.bin")?.async("uint8array"),
    ).toEqual(new Uint8Array([1, 2, 3]));

    const body = document.package.document.content.at(0);
    if (!body || body.type !== "paragraph") {
      throw new Error("expected a body paragraph");
    }
    const run = body.content.at(0);
    if (!run || run.type !== "run") {
      throw new Error("expected a body run");
    }
    const text = run.content.at(0);
    if (!text || text.type !== "text") {
      throw new Error("expected body text");
    }
    text.text = "Edited after first save";

    const secondSave = await createDocx(document);
    const secondZip = await JSZip.loadAsync(secondSave);

    expect(await secondZip.file("word/document.xml")?.async("text")).toContain(
      "Edited after first save",
    );
    expect(await secondZip.file("word/vbaProject.bin")?.async("uint8array")).toEqual(
      new Uint8Array([1, 2, 3]),
    );
  });

  test("drops orphan comment references before full repack validation", async () => {
    const originalBuffer = await createHeaderFixture();
    const document: Document = {
      originalBuffer,
      package: {
        document: {
          finalSectionProperties: {},
          content: [
            {
              type: "paragraph",
              content: [{ type: "commentReference", id: 404 }],
            },
          ],
        },
        relationships: new Map(),
      },
    };

    const result = await repackDocx(document, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(result);
    const documentXml = await zip.file("word/document.xml")?.async("text");

    expect(documentXml).not.toContain("commentReference");
    expect(documentXml).not.toContain('w:id="404"');
  });

  test("uses parser-repaired package buffers for full repack", async () => {
    const originalBuffer = await createHeaderFixture();
    const truncated = truncateAfterEndOfCentralDirectoryCounts(originalBuffer);
    const document = await parseDocx(truncated, { preloadFonts: false });

    expect(document.originalBuffer?.byteLength).toBe(originalBuffer.byteLength);

    const repacked = await repackDocx(document, { updateModifiedDate: false });
    const reparsed = await parseDocx(repacked, { preloadFonts: false });

    expect(reparsed.package.document.content).toHaveLength(1);
  });

  test("registers newly inserted header images in the header relationship part", async () => {
    const originalBuffer = await createHeaderFixture();
    const document: Document = {
      originalBuffer,
      package: {
        document: {
          finalSectionProperties: {
            headerReferences: [{ type: "default", rId: "rId1" }],
          },
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
            },
          ],
        },
        relationships: new Map([
          [
            "rId1",
            {
              id: "rId1",
              type: RELATIONSHIP_TYPES.header,
              target: "header1.xml",
            },
          ],
        ]),
        headers: new Map([
          [
            "rId1",
            {
              type: "header",
              hdrFtrType: "default",
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [
                        {
                          type: "drawing",
                          image: {
                            type: "image",
                            rId: "rId_img_123",
                            src: ONE_PIXEL_PNG_DATA_URL,
                            filename: "header.png",
                            alt: "Header image",
                            size: { width: 9525, height: 9525 },
                            wrap: { type: "inline" },
                          },
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        ]),
      },
    };

    const result = await repackDocx(document, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(result);
    const headerRels = await zip.file("word/_rels/header1.xml.rels")?.async("text");
    const headerXml = await zip.file("word/header1.xml")?.async("text");

    expect(headerRels).toContain(`Type="${RELATIONSHIP_TYPES.image}"`);
    expect(headerRels).toContain('Target="media/image1.png"');
    expect(headerXml).toContain('r:embed="rId1"');
    expect(zip.file("word/media/image1.png")).not.toBeNull();
  });

  test("does not repackage existing header images that already have relationships", async () => {
    const originalBuffer = await createHeaderFixture();
    const originalZip = await JSZip.loadAsync(originalBuffer);
    originalZip.file("word/media/image1.png", new Uint8Array([1, 2, 3, 4]));
    originalZip.file(
      "word/_rels/header1.xml.rels",
      `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
</Relationships>`,
    );
    const sourceBuffer = await originalZip.generateAsync({
      type: "arraybuffer",
    });
    const document: Document = {
      originalBuffer: sourceBuffer,
      package: {
        document: {
          finalSectionProperties: {
            headerReferences: [{ type: "default", rId: "rId1" }],
          },
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
            },
          ],
        },
        relationships: new Map([
          [
            "rId1",
            {
              id: "rId1",
              type: RELATIONSHIP_TYPES.header,
              target: "header1.xml",
            },
          ],
        ]),
        headers: new Map([
          [
            "rId1",
            {
              type: "header",
              hdrFtrType: "default",
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [
                        {
                          type: "drawing",
                          image: {
                            type: "image",
                            rId: "rId1",
                            src: ONE_PIXEL_PNG_DATA_URL,
                            filename: "image1.png",
                            size: { width: 9525, height: 9525 },
                            wrap: { type: "inline" },
                          },
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        ]),
      },
    };

    const result = await repackDocx(document, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(result);
    const headerRels = await zip.file("word/_rels/header1.xml.rels")?.async("text");
    const headerXml = await zip.file("word/header1.xml")?.async("text");
    const mediaFiles = Object.entries(zip.files)
      .filter(([path, file]) => path.startsWith("word/media/") && !file.dir)
      .map(([path]) => path);

    expect(headerRels?.match(/relationships\/image/gu)).toHaveLength(1);
    expect(headerRels).toContain('Target="media/image1.png"');
    expect(headerXml).toContain('r:embed="rId1"');
    expect(mediaFiles).toEqual(["word/media/image1.png"]);
  });

  test("preserves package-referenced images that are not browser-renderable", async () => {
    const originalBuffer = await createAlternateContentImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });

    const result = await repackDocx(doc, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(result);
    const documentXml = await zip.file("word/document.xml")?.async("text");

    expect(zip.file("word/media/image1.emf")).not.toBeNull();
    expect(documentXml).toContain("<mc:AlternateContent>");
    expect(documentXml).toContain('r:embed="rId5"');
  });

  test("full repack preserves first-section header image references", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });

    const result = await repackDocx(doc, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(result);
    const documentXml = await zip.file("word/document.xml")?.async("text");
    const headerRelsXml = await zip.file("word/_rels/header1.xml.rels")?.async("text");

    expect(documentXml).toContain('<w:headerReference w:type="first" r:id="rId11"/>');
    expect(documentXml).toContain("<w:titlePg/>");
    expect(headerRelsXml).toContain('Target="media/image1.png"');
    expect(zip.file("word/media/image1.png")).not.toBeNull();
  });

  test("fails closed when a repack would still orphan a header reference", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });
    const firstParagraph = doc.package.document.content.find(
      (block) => block.type === "paragraph" && block.sectionProperties,
    );
    if (!firstParagraph || firstParagraph.type !== "paragraph") {
      throw new Error("Expected a section-ending paragraph");
    }
    delete firstParagraph.sectionProperties;

    try {
      await repackDocx(doc, { updateModifiedDate: false });
    } catch (error) {
      expect(error).toBeInstanceOf(DocxPackageFidelityError);
      return;
    }

    throw new Error("Expected repackDocx to reject");
  });

  test("permits only a reference owned by the exactly resolved section endpoint", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });
    const firstParagraph = doc.package.document.content.find(
      (block) => block.type === "paragraph" && block.sectionProperties,
    );
    if (!firstParagraph || firstParagraph.type !== "paragraph") {
      throw new Error("Expected a section-ending paragraph");
    }
    delete firstParagraph.sectionProperties;

    const resolution = {
      type: "tracked-section-endpoint-removal",
      sourceParagraphEndpointCount: 1,
      expectedParagraphEndpointCount: 0,
      sourceEndpointFingerprint: "source-endpoints",
      expectedEndpointFingerprint: "resolved-endpoints",
      removedReferences: [{ part: "header", type: "first", relationshipId: "rId11" }],
    } as const satisfies TrackedSectionEndpointRemoval;
    const saved = await repackAfterTrackedSectionEndpointRemoval({
      document: doc,
      resolution,
    });
    const savedZip = await JSZip.loadAsync(saved);
    const savedXml = await savedZip.file("word/document.xml")?.async("text");

    expect(savedXml).not.toContain("rId11");
    expect(savedZip.file("word/header1.xml")).not.toBeNull();
  });

  test("consumes an internal section-removal allowance in one repack call", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });
    const firstParagraph = doc.package.document.content.find(
      (block) => block.type === "paragraph" && block.sectionProperties,
    );
    if (!firstParagraph || firstParagraph.type !== "paragraph") {
      throw new Error("Expected a section-ending paragraph");
    }
    delete firstParagraph.sectionProperties;

    await expect(
      withTrackedSectionEndpointRemoval({
        document: doc,
        resolution: {
          type: "tracked-section-endpoint-removal",
          sourceParagraphEndpointCount: 1,
          expectedParagraphEndpointCount: 0,
          sourceEndpointFingerprint: "source-endpoints",
          expectedEndpointFingerprint: "resolved-endpoints",
          removedReferences: [{ part: "header", type: "first", relationshipId: "rId11" }],
        },
        repack: async () => {
          await repackDocx(doc, { updateModifiedDate: false });
          return repackDocx(doc, { updateModifiedDate: false });
        },
      }),
    ).rejects.toBeInstanceOf(DocxPackageFidelityError);

    // The failed scope must not leave an allowance behind for an ordinary
    // public repack of the same document object.
    await expect(repackDocx(doc, { updateModifiedDate: false })).rejects.toBeInstanceOf(
      DocxPackageFidelityError,
    );

    // Stable tracker evidence may authorize a later save, but that save gets a
    // fresh one-call scope rather than reusing state from the failed call.
    await expect(
      repackAfterTrackedSectionEndpointRemoval({
        document: doc,
        resolution: {
          type: "tracked-section-endpoint-removal",
          sourceParagraphEndpointCount: 1,
          expectedParagraphEndpointCount: 0,
          sourceEndpointFingerprint: "source-endpoints",
          expectedEndpointFingerprint: "resolved-endpoints",
          removedReferences: [{ part: "header", type: "first", relationshipId: "rId11" }],
        },
      }),
    ).resolves.toBeInstanceOf(ArrayBuffer);
  });

  test("fails closed when serialized section loss exceeds a tracked-resolution allowance", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });
    const firstParagraph = doc.package.document.content.find(
      (block) => block.type === "paragraph" && block.sectionProperties,
    );
    if (!firstParagraph || firstParagraph.type !== "paragraph") {
      throw new Error("Expected a section-ending paragraph");
    }
    delete firstParagraph.sectionProperties;
    delete doc.package.document.finalSectionProperties;

    await expect(
      repackAfterTrackedSectionEndpointRemoval({
        document: doc,
        resolution: {
          type: "tracked-section-endpoint-removal",
          sourceParagraphEndpointCount: 1,
          expectedParagraphEndpointCount: 0,
          sourceEndpointFingerprint: "source-endpoints",
          expectedEndpointFingerprint: "resolved-endpoints",
          removedReferences: [{ part: "header", type: "first", relationshipId: "rId11" }],
        },
      }),
    ).rejects.toBeInstanceOf(DocxPackageFidelityError);
  });

  test("does not let one removed endpoint excuse a shared reference on a surviving endpoint", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const originalZip = await JSZip.loadAsync(originalBuffer);
    const originalXml = await originalZip.file("word/document.xml")?.async("text");
    if (!originalXml) {
      throw new Error("Expected document XML");
    }
    originalZip.file(
      "word/document.xml",
      originalXml.replace(
        '<w:headerReference w:type="default" r:id="rId12"/>',
        '<w:headerReference w:type="first" r:id="rId11"/>',
      ),
    );
    const sharedReferenceBuffer = await originalZip.generateAsync({ type: "arraybuffer" });
    const doc = await parseDocx(sharedReferenceBuffer, { preloadFonts: false });
    const firstParagraph = doc.package.document.content.find(
      (block) => block.type === "paragraph" && block.sectionProperties,
    );
    if (!firstParagraph || firstParagraph.type !== "paragraph") {
      throw new Error("Expected a section-ending paragraph");
    }
    delete firstParagraph.sectionProperties;
    if (!doc.package.document.finalSectionProperties) {
      throw new Error("Expected final section properties");
    }
    delete doc.package.document.finalSectionProperties.headerReferences;

    await expect(
      repackAfterTrackedSectionEndpointRemoval({
        document: doc,
        resolution: {
          type: "tracked-section-endpoint-removal",
          sourceParagraphEndpointCount: 1,
          expectedParagraphEndpointCount: 0,
          sourceEndpointFingerprint: "source-endpoints",
          expectedEndpointFingerprint: "resolved-endpoints",
          removedReferences: [{ part: "header", type: "first", relationshipId: "rId11" }],
        },
      }),
    ).rejects.toBeInstanceOf(DocxPackageFidelityError);
  });

  test("selective save preserves first-section header image references", async () => {
    const originalBuffer = await createMultiSectionFirstHeaderImageFixture();
    const doc = await parseDocx(originalBuffer, { preloadFonts: false });

    const result = await attemptSelectiveSave(doc, originalBuffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("Expected selective save result");
    }

    const zip = await JSZip.loadAsync(result);
    const documentXml = await zip.file("word/document.xml")?.async("text");
    const headerRelsXml = await zip.file("word/_rels/header1.xml.rels")?.async("text");

    expect(documentXml).toContain('<w:headerReference w:type="first" r:id="rId11"/>');
    expect(documentXml).toContain("<w:titlePg/>");
    expect(headerRelsXml).toContain('Target="media/image1.png"');
    expect(zip.file("word/media/image1.png")).not.toBeNull();
  });
});
