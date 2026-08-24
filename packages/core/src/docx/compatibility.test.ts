import { describe, expect, test } from "bun:test";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, DrawingContent, ShapeContent } from "../types/document";
import { inspectDocxCompatibility } from "./compatibility";
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
};

const createDocument = ({ imageSrc, paraId, rawXml }: CreateDocumentOptions = {}): Document => ({
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
                    rId: "rId1",
                    ...(imageSrc === undefined ? {} : { src: imageSrc }),
                    size: { width: 9525, height: 9525 },
                    wrap: { type: "inline" },
                  },
                  ...(rawXml !== undefined ? { rawXml } : {}),
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
      schemaVersion: 1,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: true,
      issues: [],
      reasons: [],
      unsupportedContentCount: 0,
    });
  });

  test("reports a renderable raw drawing at its document location", () => {
    expect(
      inspectDocxCompatibility(
        createDocument({
          imageSrc: "data:image/png;base64,AA==",
          paraId: "A1B2C3D4",
          rawXml: "<w:drawing/>",
        }),
      ),
    ).toEqual({
      schemaVersion: 1,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: false,
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

  test("reports unresolved raw media without a preservation-only classification", () => {
    expect(
      inspectDocxCompatibility(
        createDocument({
          paraId: "A1B2C3D4",
          rawXml: '<w:drawing><a:blip r:embed="rId1"/></w:drawing>',
        }),
      ).canSafelyEdit,
    ).toBe(false);
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

  test("uses parsed package metadata unless the caller overrides it", () => {
    const document = createDocument({
      imageSrc: "data:image/png;base64,AA==",
      rawXml: "<w:drawing/>",
    });
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
