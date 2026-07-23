import { describe, expect, test } from "bun:test";

import type { Document, Image, Run } from "../../types/document";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { parseDrawing } from "../imageParser";
import { serializeRun } from "../serializer/runSerializer";
import { parseXml, type XmlElement } from "../xmlParser";

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

type DrawingKind = "anchor" | "inline";

const drawingXml = (kind: DrawingKind, docPrAttrs: string): string => {
  const open =
    kind === "inline"
      ? "<wp:inline>"
      : '<wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">';
  const position =
    kind === "anchor"
      ? '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
      : "";
  const wrap = kind === "anchor" ? '<wp:wrapSquare wrapText="bothSides"/>' : "";

  return `<w:drawing ${NS}>
    ${open}
      ${position}
      <wp:extent cx="914400" cy="457200"/>
      ${wrap}
      <wp:docPr id="7" ${docPrAttrs}/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="7" name="media.png"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:${kind}>
  </w:drawing>`;
};

const parseDrawingXml = (xml: string): Image | null => {
  const parsed = parseXml(xml);
  const drawing = (parsed.elements as XmlElement[]).at(0);
  return drawing ? parseDrawing(drawing, undefined, undefined) : null;
};

const serializeImage = (image: Image): string => {
  const run: Run = { type: "run", content: [{ type: "drawing", image }] };
  return serializeRun(run);
};

const reparseSerializedImage = (xml: string): Image | null => {
  const parsed = parseXml(`<root ${NS}>${xml}</root>`);
  const root = (parsed.elements as XmlElement[]).at(0);
  const run = (root?.elements as XmlElement[] | undefined)?.at(0);
  const drawing = (run?.elements as XmlElement[] | undefined)?.at(0);
  return drawing ? parseDrawing(drawing, undefined, undefined) : null;
};

const documentWithImage = (image: Image): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "drawing", image }] }],
        },
      ],
    },
  },
});

const firstImage = (document: Document): Image | null => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    return null;
  }
  const run = paragraph.content.at(0);
  if (run?.type !== "run") {
    return null;
  }
  const drawing = run.content.at(0);
  return drawing?.type === "drawing" ? drawing.image : null;
};

describe("wp:docPr image metadata round-trip", () => {
  test.each(["inline", "anchor"] as const)(
    "keeps authored name, description, and title distinct for %s images",
    (kind) => {
      const image = parseDrawingXml(
        drawingXml(
          kind,
          'name="Authored name" descr="Accessibility description" title="Display title"',
        ),
      );
      expect(image).toMatchObject({
        docPrName: "Authored name",
        alt: "Accessibility description",
        title: "Display title",
      });
      if (!image) {
        return;
      }

      const serialized = serializeImage(image);
      expect(serialized).toContain(
        'name="Authored name" descr="Accessibility description" title="Display title"',
      );
      expect(reparseSerializedImage(serialized)).toMatchObject({
        docPrName: "Authored name",
        alt: "Accessibility description",
        title: "Display title",
      });

      const editorRoundTrip = fromProseDoc(
        toProseDoc(documentWithImage(image)),
        documentWithImage(image),
      );
      expect(firstImage(editorRoundTrip)).toMatchObject({
        docPrName: "Authored name",
        alt: "Accessibility description",
        title: "Display title",
      });
    },
  );

  test("does not synthesize title from name or description", () => {
    const image = parseDrawingXml(
      drawingXml("inline", 'name="Authored name" descr="Accessibility description"'),
    );
    expect(image?.title).toBeUndefined();
    if (!image) {
      return;
    }

    const serialized = serializeImage(image);
    expect(serialized).toContain('name="Authored name" descr="Accessibility description"');
    expect(serialized).not.toContain(" title=");
    expect(reparseSerializedImage(serialized)?.title).toBeUndefined();
  });
});
