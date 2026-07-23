// Regression eigenpal #567: Word stores anchored wps:wsp text boxes inside
// an <mc:AlternateContent> block (Choice Requires="wps" + Fallback VML).
// scanRunForTextBoxDrawings only walked direct children of <w:r>, so the
// <w:drawing> inside the Choice branch never reached the text-box pipeline
// and the shape text (e.g. "Organisation Chart" cards) was silently dropped.

import { describe, expect, test } from "bun:test";

import type { MediaFile, RelationshipMap } from "../types/document";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocumentBody } from "./documentParser";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const textBoxDrawingXml = (text: string) => `
  <w:drawing>
    <wp:inline>
      <wp:extent cx="914400" cy="457200"/>
      <wp:docPr id="11" name="Text Box 11"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;

describe("parseDocumentBody — AlternateContent text boxes", () => {
  test("parses theme-coloured straight connectors as editable shapes", () => {
    const body = parseDocumentBody(`${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body><w:p><w:r><mc:AlternateContent>
    <mc:Choice Requires="wps"><w:drawing><wp:anchor>
      <wp:extent cx="0" cy="25400"/>
      <wp:positionH relativeFrom="column"><wp:posOffset>1</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>
      <wp:wrapNone/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp><wps:spPr>
          <a:xfrm><a:off x="2323400" y="3780000"/><a:ext cx="6045200" cy="0"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln w="25400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>
        </wps:spPr></wps:wsp>
      </a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice>
    <mc:Fallback><w:drawing/></mc:Fallback>
  </mc:AlternateContent></w:r></w:p></w:body>
</w:document>`);

    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }
    const shape = paragraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content)
      .find((content) => content.type === "shape");

    expect(shape?.type).toBe("shape");
    if (shape?.type !== "shape") {
      throw new Error("Expected editable shape");
    }
    expect(shape.shape).toMatchObject({
      shapeType: "straightConnector1",
      size: { width: 6_045_200, height: 25_400 },
      outline: {
        color: { themeColor: "accent1" },
        width: 25_400,
      },
    });

    const converted = fromProseDoc(toProseDoc({ package: { document: body } }));
    const convertedParagraph = converted.package.document.content.at(0);
    if (convertedParagraph?.type !== "paragraph") {
      throw new Error("Expected converted paragraph");
    }
    const convertedShape = convertedParagraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content)
      .find((content) => content.type === "shape");
    expect(
      convertedShape?.type === "shape" ? convertedShape.shape.outline?.color : undefined,
    ).toEqual({ themeColor: "accent1" });
  });

  test("prefers a renderable grouped Choice over a flattened fallback picture", () => {
    const rels: RelationshipMap = new Map([
      [
        "rIdChoice",
        {
          id: "rIdChoice",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: "media/choice.png",
        },
      ],
      [
        "rIdFallback",
        {
          id: "rIdFallback",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: "media/fallback.png",
        },
      ],
    ]);
    const media = new Map<string, MediaFile>([
      [
        "word/media/choice.png",
        {
          path: "word/media/choice.png",
          mimeType: "image/png",
          data: new ArrayBuffer(0),
          dataUrl: "data:image/png;base64,Y2hvaWNl",
        },
      ],
      [
        "word/media/fallback.png",
        {
          path: "word/media/fallback.png",
          mimeType: "image/png",
          data: new ArrayBuffer(0),
          dataUrl: "data:image/png;base64,ZmFsbGJhY2s=",
        },
      ],
    ]);
    const body = parseDocumentBody(
      `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body><w:p><w:r><mc:AlternateContent>
    <mc:Choice Requires="wpg"><w:drawing><wp:anchor behindDoc="1">
      <wp:extent cx="2000000" cy="1000000"/><wp:wrapNone/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp>
        <pic:pic><pic:blipFill><a:blip r:embed="rIdChoice"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></pic:spPr></pic:pic>
      </wpg:wgp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice>
    <mc:Fallback><w:pict><v:shape style="width:9000;height:4000"><v:imagedata r:id="rIdFallback"/></v:shape></w:pict></mc:Fallback>
  </mc:AlternateContent></w:r></w:p></w:body>
</w:document>`,
      null,
      null,
      null,
      rels,
      media,
    );

    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }
    const drawing = paragraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content)
      .find((content) => content.type === "drawing");
    expect(drawing?.type).toBe("drawing");
    if (drawing?.type !== "drawing") {
      throw new Error("Expected drawing");
    }
    expect(drawing.image.mimeType).toBe("image/svg+xml");
    expect(drawing.image.size).toEqual({ width: 2_000_000, height: 1_000_000 });
    expect(decodeURIComponent(drawing.image.src ?? "")).toContain("Y2hvaWNl");
    expect(decodeURIComponent(drawing.image.src ?? "")).not.toContain("ZmFsbGJhY2s=");
    expect(drawing.rawXml).toContain("<mc:Fallback>");
  });

  test("renders a wpg Choice as SVG while preserving the complete alternate content", () => {
    const body = parseDocumentBody(`${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body><w:p><w:r><mc:AlternateContent>
    <mc:Choice Requires="wpg"><w:drawing><wp:anchor behindDoc="1">
      <wp:extent cx="1000000" cy="500000"/><wp:wrapTopAndBottom/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp>
        <wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="DBEDF3"/></a:solidFill></wps:spPr></wps:wsp>
      </wpg:wgp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></mc:Choice>
    <mc:Fallback><w:pict/></mc:Fallback>
  </mc:AlternateContent></w:r></w:p></w:body>
</w:document>`);

    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }
    const drawing = paragraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content)
      .find((content) => content.type === "drawing");

    expect(drawing?.type).toBe("drawing");
    if (drawing?.type !== "drawing") {
      throw new Error("Expected grouped drawing");
    }
    expect(drawing.image.mimeType).toBe("image/svg+xml");
    expect(drawing.image.src).toStartWith("data:image/svg+xml");
    expect(drawing.rawXml).toContain("<mc:AlternateContent");
    expect(drawing.rawXml).toContain("<mc:Fallback>");
  });

  test("extracts text-box drawings wrapped in mc:AlternateContent", () => {
    const body = parseDocumentBody(`${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:p w14:paraId="ACTB0001">
      <w:r><w:t>Host text</w:t></w:r>
      <w:r>
        <mc:AlternateContent>
          <mc:Choice Requires="wps">${textBoxDrawingXml("Card title")}</mc:Choice>
          <mc:Fallback>
            <w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>VML fallback</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>
  </w:body>
</w:document>`);

    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    const shapes = paragraph.content
      .filter((c) => c.type === "run")
      .flatMap((r) => r.content.filter((c) => c.type === "shape"));

    expect(shapes).toHaveLength(1);
    const shape = shapes.at(0);
    if (!shape) {
      throw new Error("Expected shape");
    }
    expect(shape.shape.shapeType).toBe("textBox");
    const innerPara = shape.shape.textBody?.content.at(0);
    if (innerPara?.type !== "paragraph") {
      throw new Error("Expected text-box inner paragraph");
    }
    const innerRun = innerPara.content.at(0);
    if (innerRun?.type !== "run") {
      throw new Error("Expected text-box inner run");
    }
    const innerText = innerRun.content.at(0);
    if (innerText?.type !== "text") {
      throw new Error("Expected text-box inner text");
    }
    expect(innerText.text).toBe("Card title");
  });

  // Regression: with multiple AlternateContent-only <w:r> elements followed by
  // a text run, a strict index check could drop shapes past the first (since
  // consolidateParagraphContent collapses shape-only <w:r> into surrounding
  // parsed runs). All shapes must survive.
  test("preserves multiple AlternateContent-only runs in the same paragraph", () => {
    const body = parseDocumentBody(`${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:p w14:paraId="ACTB0002">
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Card 1")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Card 2")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Card 3")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
      <w:r><w:t>Trailing text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`);

    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    const cardTitles = paragraph.content
      .filter((c) => c.type === "run")
      .flatMap((r) => r.content.filter((c) => c.type === "shape"))
      .map((shape) => {
        const inner = shape.shape.textBody?.content.at(0);
        if (inner?.type !== "paragraph") {
          return null;
        }
        const run = inner.content.at(0);
        if (run?.type !== "run") {
          return null;
        }
        const text = run.content.at(0);
        return text?.type === "text" ? text.text : null;
      });

    expect(cardTitles).toEqual(["Card 1", "Card 2", "Card 3"]);
  });

  test("preserves text boxes when the host paragraph has no flow content", () => {
    const body = parseDocumentBody(`${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:p>
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Card 1")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Card 2")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
    </w:p>
  </w:body>
</w:document>`);

    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    const shapes = paragraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content.filter((content) => content.type === "shape"));

    expect(shapes).toHaveLength(2);
  });

  test("preserves text boxes anchored from a table-cell paragraph", () => {
    const body = parseDocumentBody(`${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:tbl><w:tr><w:tc><w:p>
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Cell card 1")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
      <w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBoxDrawingXml("Cell card 2")}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
    </w:p></w:tc></w:tr></w:tbl>
  </w:body>
</w:document>`);

    const table = body.content.at(0);
    if (table?.type !== "table") {
      throw new Error("Expected table");
    }
    const paragraph = table.rows.at(0)?.cells.at(0)?.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected table-cell paragraph");
    }

    const cardTitles = paragraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content.filter((content) => content.type === "shape"))
      .map((shape) => {
        const innerParagraph = shape.shape.textBody?.content.at(0);
        if (innerParagraph?.type !== "paragraph") {
          return null;
        }
        const innerRun = innerParagraph.content.at(0);
        if (innerRun?.type !== "run") {
          return null;
        }
        const innerText = innerRun.content.at(0);
        return innerText?.type === "text" ? innerText.text : null;
      });

    expect(cardTitles).toEqual(["Cell card 1", "Cell card 2"]);
  });
});
