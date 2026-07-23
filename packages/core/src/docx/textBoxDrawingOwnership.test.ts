import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const IMAGE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const textBoxDrawing = (text: string): string => `
  <w:drawing>
    <wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="1828800" cy="914400"/>
      <wp:wrapSquare wrapText="bothSides"/>
      <wp:docPr id="1" name="Text box"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>`;

const documentXml = (runContent: string): string => `${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body><w:p><w:r>${runContent}</w:r></w:p><w:sectPr/></w:body>
</w:document>`;

const buildDocx = async (runContent: string, includeImage = false): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentXml(runContent));

  if (includeImage) {
    const relsPath = "word/_rels/document.xml.rels";
    const relationships = await zip.file(relsPath)!.async("text");
    zip.file(
      relsPath,
      relationships.replace(
        "</Relationships>",
        `<Relationship Id="rIdImage" Type="${IMAGE_RELATIONSHIP_TYPE}" Target="media/preview.png"/></Relationships>`,
      ),
    );
    zip.file("word/media/preview.png", new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  }

  return zip.generateAsync({ type: "arraybuffer" });
};

type ParsedDocx = Awaited<ReturnType<typeof parseDocx>>;

const bodyDrawingTypes = ({ package: { document } }: ParsedDocx): string[] =>
  document.content.flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((content) =>
          content.type === "run" ? content.content.map(({ type }) => type) : [],
        )
      : [],
  );

const savedDocumentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file("word/document.xml")!.async("text");
};

describe("text-box drawing ownership", () => {
  test("a modern AlternateContent text box owns its VML fallback exactly once", async () => {
    const source = await buildDocx(`
      <mc:AlternateContent>
        <mc:Choice Requires="wps">${textBoxDrawing("Editable text")}</mc:Choice>
        <mc:Fallback>
          <w:pict>
            <v:rect style="width:2in;height:1in">
              <v:textbox><w:txbxContent><w:p><w:r><w:t>Fallback text</w:t></w:r></w:p></w:txbxContent></v:textbox>
            </v:rect>
          </w:pict>
        </mc:Fallback>
      </mc:AlternateContent>`);
    const parsed = await parseDocx(source, { preloadFonts: false });

    expect(bodyDrawingTypes(parsed)).toEqual(["shape"]);

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(bodyDrawingTypes(reopened)).toEqual(["shape"]);
    expect((await savedDocumentXml(saved)).match(/<w:drawing(?:\s|>)/gu)).toHaveLength(1);
  });

  test("an image-backed VML text-box container stays one raw drawing", async () => {
    const source = await buildDocx(
      `<w:pict>
        <v:shape style="width:2in;height:1in">
          <v:imagedata r:id="rIdImage"/>
          <v:textbox><w:txbxContent><w:p><w:r><w:t>Legacy overlay</w:t></w:r></w:p></w:txbxContent></v:textbox>
        </v:shape>
      </w:pict>`,
      true,
    );
    const parsed = await parseDocx(source, { preloadFonts: false });

    expect(bodyDrawingTypes(parsed)).toEqual(["drawing"]);

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(bodyDrawingTypes(reopened)).toEqual(["drawing"]);
    const savedXml = await savedDocumentXml(saved);
    expect(savedXml.match(/<w:pict(?:\s|>)/gu)).toHaveLength(1);
    expect(savedXml).not.toContain("<w:drawing");
  });
});
