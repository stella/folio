import { describe, expect, test } from "bun:test";
import { DOCX_CONFORMANCE_CLASSES, DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, DrawingContent, ShapeContent } from "../types/document";
import { inspectDocxCompatibility } from "./compatibility";
import { parseDocumentBody } from "./documentParser";
import { classifyDrawingSafety } from "./imageRawXml";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const createAdjustedShapeDocx = async (
  preset: "rightBrace" | "roundRect",
): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  const adjustments =
    preset === "rightBrace"
      ? '<a:gd name="adj1" fmla="val 12500"/><a:gd name="adj2" fmla="val 62500"/>'
      : '<a:gd name="adj" fmla="val 25000"/>';
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
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
    "word/document.xml",
    `${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p><w:r><mc:AlternateContent>
      <mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="914400" cy="457200"/><wp:wrapNone/>
        <wp:docPr id="1" name="Adjusted shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
            <a:prstGeom prst="${preset}"><a:avLst>${adjustments}</a:avLst></a:prstGeom>
            <a:noFill/><a:ln><a:solidFill><a:srgbClr val="666666"/></a:solidFill></a:ln>
          </wps:spPr></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:anchor></w:drawing></mc:Choice>
      <mc:Fallback><w:pict><v:shape id="fallback-shape"/></w:pict></mc:Fallback>
    </mc:AlternateContent></w:r></w:p>
    <w:p><w:r><w:t>Editable text</w:t></w:r></w:p>
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

/** Smallest valid PNG, so the header logo resolves to real media. */
const ONE_PIXEL_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const HEADER_PICTURE_XML = `<w:p><w:r><w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="190500" cy="190500"/>
    <wp:effectExtent l="12700" t="19050" r="6350" b="25400"/>
    <wp:docPr id="7" name="Logo" descr="Mark"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="0" noMove="1"/></wp:cNvGraphicFramePr>
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

const CELL_SHAPE_XML = `<w:r><w:drawing><wp:anchor behindDoc="0" layoutInCell="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
  <wp:extent cx="114300" cy="457200"/><wp:wrapNone/>
  <wp:docPr id="3" name="Brace"/>
  <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
    <wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="114300" cy="457200"/></a:xfrm>
      <a:prstGeom prst="rightBrace"><a:avLst>
        <a:gd name="adj1" fmla="val 12500"/><a:gd name="adj2" fmla="val 62500"/>
      </a:avLst></a:prstGeom>
      <a:noFill/><a:ln><a:solidFill><a:srgbClr val="666666"/></a:solidFill></a:ln>
    </wps:spPr></wps:wsp>
  </a:graphicData></a:graphic>
</wp:anchor></w:drawing></w:r>`;

const XML_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");

/** A header logo plus a bodiless brace shape in a table cell: no opaque content. */
const createHeaderPictureDocx = async (): Promise<ArrayBuffer> => {
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
  <Relationship Id="rId10" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/header1.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.image}" Target="media/image1.png"/>
</Relationships>`,
  );
  zip.file("word/media/image1.png", ONE_PIXEL_PNG_BASE64, { base64: true });
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}
<w:hdr ${XML_NAMESPACES}>${HEADER_PICTURE_XML}</w:hdr>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document ${XML_NAMESPACES}>
  <w:body>
    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>
        <w:p>${CELL_SHAPE_XML}</w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
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

/** A `wpg:wgp` group the rasterizer renders into a single preview image. */
const GROUPED_DRAWING_BODY_XML = `${XML_DECLARATION}
<w:document ${XML_NAMESPACES}>
  <w:body><w:p><w:r><w:drawing><wp:anchor behindDoc="1">
    <wp:extent cx="1000000" cy="500000"/><wp:wrapTopAndBottom/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp>
      <wps:wsp><wps:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="250000"/></a:xfrm>
        <a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="DBEDF3"/></a:solidFill>
      </wps:spPr></wps:wsp>
      <pic:pic>
        <pic:blipFill><a:blip r:embed="rIdGroupChild"/></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="250000"/><a:ext cx="1000000" cy="250000"/></a:xfrm></pic:spPr>
      </pic:pic>
    </wpg:wgp></a:graphicData></a:graphic>
  </wp:anchor></w:drawing></w:r></w:p></w:body>
</w:document>`;

const firstRawDrawing = (document: Document): DrawingContent | undefined => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    return undefined;
  }

  return paragraph.content
    .filter((item) => item.type === "run")
    .flatMap((run) => run.content)
    .find(
      (content): content is DrawingContent =>
        content.type === "drawing" && content.rawXml !== undefined,
    );
};

const firstShape = (document: Document): ShapeContent | undefined => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    return undefined;
  }

  return paragraph.content
    .filter((item) => item.type === "run")
    .flatMap((run) => run.content)
    .find((content): content is ShapeContent => content.type === "shape");
};

type CreateDocumentOptions = {
  imageSrc?: string;
  paraId?: string;
  rawXml?: string;
  /** Empty models a blip whose media relationship never resolved. */
  rId?: string;
  /** A fingerprint that no longer matches the image makes the raw XML stale. */
  staleFingerprint?: boolean;
};

const createDocument = ({
  imageSrc,
  paraId,
  rawXml,
  rId = "rId1",
  staleFingerprint = false,
}: CreateDocumentOptions = {}): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          ...(paraId === undefined ? {} : { paraId }),
          content: [
            {
              type: "run",
              content: [
                { type: "text", text: "Body" },
                {
                  type: "drawing",
                  image: {
                    type: "image",
                    rId,
                    ...(imageSrc === undefined ? {} : { src: imageSrc }),
                    size: { width: 9525, height: 9525 },
                    wrap: { type: "inline" },
                  },
                  ...(rawXml !== undefined ? { rawXml } : {}),
                  ...(staleFingerprint ? { rawImageFingerprint: "stale" } : {}),
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

describe("DOCX compatibility inspection", () => {
  test("allows editing ordinary parsed content", () => {
    expect(inspectDocxCompatibility(createDocument())).toEqual({
      schemaVersion: 2,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: true,
      drawings: [
        {
          class: "native",
          location: {
            part: { type: "document" },
            path: "package.document.content[0].content[0].content[1]",
          },
        },
      ],
      issues: [],
      reasons: [],
      unsupportedContentCount: 0,
    });
  });

  test("keeps a picture whose raw XML the serializer replays editable", () => {
    expect(
      inspectDocxCompatibility(
        createDocument({
          imageSrc: "data:image/png;base64,AA==",
          paraId: "A1B2C3D4",
          rawXml: "<w:drawing/>",
        }),
      ),
    ).toEqual({
      schemaVersion: 2,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: true,
      drawings: [
        {
          class: "replayable",
          location: {
            blockId: "A1B2C3D4",
            part: { type: "document" },
            path: "package.document.content[0].content[0].content[1]",
          },
        },
      ],
      issues: [],
      reasons: [],
      unsupportedContentCount: 0,
    });
  });

  test("blocks raw media the serializer can neither replay nor regenerate", () => {
    expect(
      inspectDocxCompatibility(
        createDocument({
          paraId: "A1B2C3D4",
          rawXml: '<w:drawing><a:blip r:embed="rId1"/></w:drawing>',
          rId: "",
          staleFingerprint: true,
        }),
      ),
    ).toEqual({
      schemaVersion: 2,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: false,
      drawings: [
        {
          class: "opaque",
          location: {
            blockId: "A1B2C3D4",
            part: { type: "document" },
            path: "package.document.content[0].content[0].content[1]",
          },
        },
      ],
      issues: [
        {
          code: "opaqueDrawing",
          location: {
            blockId: "A1B2C3D4",
            part: { type: "document" },
            path: "package.document.content[0].content[0].content[1]",
          },
        },
      ],
      reasons: ["opaqueDrawing"],
      unsupportedContentCount: 1,
    });
  });

  test("keeps a stale raw picture editable while its relationship still regenerates", () => {
    const compatibility = inspectDocxCompatibility(
      createDocument({
        imageSrc: "data:image/png;base64,AA==",
        rawXml: "<w:drawing/>",
        staleFingerprint: true,
      }),
    );

    expect(compatibility.drawings.at(0)?.class).toBe("native");
    expect(compatibility.canSafelyEdit).toBe(true);
  });

  test("allows text edits while preserving an unmodeled drawing verbatim", async () => {
    const source = await createAdjustedShapeDocx("roundRect");
    const parsed = await parseDocx(source, { detectVariables: false, preloadFonts: false });
    const originalDrawing = firstRawDrawing(parsed);

    expect(originalDrawing?.rawXml).toContain("<mc:AlternateContent");
    expect(originalDrawing?.rawXmlMode).toBe("preserveOnly");
    expect(originalDrawing?.image.src).toBeUndefined();
    expect(inspectDocxCompatibility(parsed).canSafelyEdit).toBe(true);

    const pmDocument = toProseDoc(parsed);
    let editableTextPosition: number | undefined;
    pmDocument.descendants((node, position) => {
      if (node.isText && node.text === "Editable text") {
        editableTextPosition = position;
      }
    });
    if (editableTextPosition === undefined) {
      throw new Error("Expected editable fixture text");
    }

    const state = EditorState.create({ doc: pmDocument });
    const editedPmDocument = state.apply(
      state.tr.insertText(
        "Updated text",
        editableTextPosition,
        editableTextPosition + "Editable text".length,
      ),
    ).doc;
    const saved = await repackDocx(fromProseDoc(editedPmDocument, parsed), {
      updateModifiedDate: false,
    });
    const reopened = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    const reopenedDrawing = firstRawDrawing(reopened);

    expect(toProseDoc(reopened).textContent).toContain("Updated text");
    expect(reopenedDrawing?.rawXml).toBe(originalDrawing?.rawXml);
    expect(inspectDocxCompatibility(reopened).canSafelyEdit).toBe(true);
  });

  test("keeps a document editable when its only drawings are a header logo and a cell shape", async () => {
    const parsed = await parseDocx(await createHeaderPictureDocx(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const compatibility = inspectDocxCompatibility(parsed);

    expect(compatibility.drawings).toEqual([
      {
        class: "replayable",
        location: {
          part: { type: "header", relationshipId: "rId10" },
          path: 'package.headers.get("rId10").content[0].content[0].content[0]',
        },
      },
    ]);
    expect(compatibility.canSafelyEdit).toBe(true);
  });

  test("replays an untouched header verbatim through a body-only edit", async () => {
    const source = await createHeaderPictureDocx();
    const parsed = await parseDocx(source, { detectVariables: false, preloadFonts: false });
    const originalHeaderXml = await (
      await JSZip.loadAsync(source)
    )
      .file("word/header1.xml")
      ?.async("text");

    const pmDocument = toProseDoc(parsed);
    let bodyTextPosition: number | undefined;
    pmDocument.descendants((node, position) => {
      if (node.isText && node.text === "Body text") {
        bodyTextPosition = position;
      }
    });
    if (bodyTextPosition === undefined) {
      throw new Error("Expected editable fixture text");
    }

    const state = EditorState.create({ doc: pmDocument });
    const edited = state.apply(
      state.tr.insertText("Edited body", bodyTextPosition, bodyTextPosition + "Body text".length),
    ).doc;
    const saved = await repackDocx(fromProseDoc(edited, parsed), { updateModifiedDate: false });
    const savedHeaderXml = await (
      await JSZip.loadAsync(saved)
    )
      .file("word/header1.xml")
      ?.async("text");

    expect(savedHeaderXml).toBe(originalHeaderXml);
    expect(
      toProseDoc(await parseDocx(saved, { detectVariables: false, preloadFonts: false }))
        .textContent,
    ).toContain("Edited body");
  });

  test("blocks on opaque content in every part the save re-serializes", () => {
    const opaqueRun = createDocument({ rawXml: "<w:drawing/>", rId: "", staleFingerprint: true })
      .package.document.content;

    for (const document of [
      {
        package: {
          document: { content: opaqueRun },
        },
      },
      {
        package: {
          document: { content: [] },
          headers: new Map([
            ["rId7", { type: "header", hdrFtrType: "default", content: opaqueRun }],
          ]),
        },
      },
      {
        package: {
          document: { content: [] },
          footers: new Map([
            ["rId8", { type: "footer", hdrFtrType: "default", content: opaqueRun }],
          ]),
        },
      },
      {
        package: {
          document: { content: [] },
          footnotes: [{ id: 2, content: opaqueRun }],
        },
      },
      {
        package: {
          document: { content: [] },
          endnotes: [{ id: 3, content: opaqueRun }],
        },
      },
    ] satisfies Document[]) {
      expect(inspectDocxCompatibility(document).canSafelyEdit).toBe(false);
    }
  });

  test("treats a supported adjusted right brace as editable shape content", async () => {
    const source = await createAdjustedShapeDocx("rightBrace");
    const parsed = await parseDocx(source, { detectVariables: false, preloadFonts: false });
    const expectedAdjustments = [
      { name: "adj1", formula: "val 12500" },
      { name: "adj2", formula: "val 62500" },
    ];

    expect(firstShape(parsed)?.shape).toMatchObject({
      shapeType: "rightBrace",
      geometryAdjustments: expectedAdjustments,
    });
    expect(firstRawDrawing(parsed)).toBeUndefined();
    expect(inspectDocxCompatibility(parsed).canSafelyEdit).toBe(true);

    const pmRoundTripped = fromProseDoc(toProseDoc(parsed), parsed);
    expect(firstShape(pmRoundTripped)?.shape.geometryAdjustments).toEqual(expectedAdjustments);

    const saved = await repackDocx(pmRoundTripped, { updateModifiedDate: false });
    const reopened = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    expect(firstShape(reopened)?.shape).toMatchObject({
      shapeType: "rightBrace",
      geometryAdjustments: expectedAdjustments,
    });
    expect(inspectDocxCompatibility(reopened).canSafelyEdit).toBe(true);
  });

  test("reports the requested profile, host, and non-body part", () => {
    const document = createDocument({
      imageSrc: "data:image/png;base64,AA==",
      paraId: "A1B2C3D4",
      rawXml: "<w:drawing/>",
      rId: "",
      staleFingerprint: true,
    });
    const content = document.package.document.content;
    document.package.document.content = [];
    document.package.headers = new Map([
      [
        "rId7",
        {
          type: "header",
          hdrFtrType: "default",
          content,
        },
      ],
    ]);

    const compatibility = inspectDocxCompatibility(document, {
      host: "browser",
      profile: "transitional",
    });

    expect(compatibility.context).toEqual({ host: "browser", profile: "transitional" });
    expect(compatibility.issues.at(0)?.location).toEqual({
      blockId: "A1B2C3D4",
      part: { type: "header", relationshipId: "rId7" },
      path: 'package.headers.get("rId7").content[0].content[0].content[1]',
    });
  });

  test("blocks a rasterized group once its preview no longer matches the raw XML", () => {
    const body = parseDocumentBody(GROUPED_DRAWING_BODY_XML);
    const document: Document = { package: { document: body } };
    const drawing = firstRawDrawing(document);
    if (drawing === undefined) {
      throw new Error("Expected the grouped fixture to parse as a raw drawing");
    }

    expect(drawing.rawXmlMode).toBe(DRAWING_RAW_XML_MODES.PREVIEW_ONLY);
    expect(classifyDrawingSafety(drawing)).toBe("replayable");
    expect(inspectDocxCompatibility(document).canSafelyEdit).toBe(true);

    // Editing the preview is editing a picture of the group, so the save can
    // neither replay the raw XML nor regenerate the group from the model.
    drawing.image.size.width += 10;

    expect(classifyDrawingSafety(drawing)).toBe("opaque");
    const compatibility = inspectDocxCompatibility(document);
    expect(compatibility.canSafelyEdit).toBe(false);
    expect(compatibility.reasons).toEqual(["opaqueDrawing"]);
    expect(compatibility.issues).toEqual([
      {
        code: "opaqueDrawing",
        location: {
          part: { type: "document" },
          path: "package.document.content[0].content[0].content[0]",
        },
      },
    ]);

    // `parseGroupDrawing` leaves the preview relationship-less today, which on
    // its own would classify a stale group as opaque. Bind one so the verdict
    // can only come from the preview-only mode: whether the rasterizer ever
    // resolves a child blip must not decide whether a group is saveable.
    drawing.image.rId = "rIdGroupChild";

    expect(classifyDrawingSafety(drawing)).toBe("opaque");
  });

  test("uses parsed package metadata unless the caller overrides it", () => {
    const document = createDocument({ imageSrc: "data:image/png;base64,AA==" });
    document.package.conformanceClass = DOCX_CONFORMANCE_CLASSES.STRICT;

    expect(inspectDocxCompatibility(document).context.profile).toBe(
      DOCX_CONFORMANCE_CLASSES.STRICT,
    );
    expect(
      inspectDocxCompatibility(document, {
        profile: DOCX_CONFORMANCE_CLASSES.TRANSITIONAL,
      }).context.profile,
    ).toBe(DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
  });
});
