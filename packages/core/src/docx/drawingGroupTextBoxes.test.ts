import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, DrawingContent, Paragraph, Shape, ShapeContent } from "../types/document";
import { parseDocumentBody } from "./documentParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");

const textBoxShape = (
  name: string,
  frame: { x: number; y: number; cx: number; cy: number },
  text: string,
  bodyPr = '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>',
): string =>
  `<wps:wsp><wps:cNvPr id="9" name="${name}"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="${frame.x}" y="${frame.y}"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>${bodyPr}</wps:wsp>`;

const rectangle = (frame: { x: number; y: number; cx: number; cy: number }): string =>
  `<wps:wsp><wps:cNvPr id="8" name="Fill"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="${frame.x}" y="${frame.y}"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D9E1F3"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>`;

/**
 * A floating group 2,000,000 x 1,000,000 EMU at (500,000, 250,000), whose
 * children are authored in a 1,000 x 500 unit space starting at (100, -200):
 * one child unit spans 2,000 EMU on both axes.
 */
const anchoredGroup = (children: string, placement?: string): string =>
  `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="7" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${
    placement ??
    '<wp:positionH relativeFrom="page"><wp:posOffset>500000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>250000</wp:posOffset></wp:positionV>'
  }<wp:extent cx="2000000" cy="1000000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="3" name="Group 3"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="100" y="-200"/><a:chExt cx="1000" cy="500"/></a:xfrm></wpg:grpSpPr>${children}</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;

const TWO_TEXT_BOXES =
  rectangle({ x: 100, y: -200, cx: 1000, cy: 500 }) +
  textBoxShape("First", { x: 150, y: -150, cx: 400, cy: 100 }, "Alpha") +
  textBoxShape(
    "Second",
    { x: 600, y: 0, cx: 300, cy: 200 },
    "Beta",
    '<wps:bodyPr lIns="45720" tIns="0" rIns="45720" bIns="0" anchor="ctr"><a:noAutofit/></wps:bodyPr>',
  );

const NESTED =
  rectangle({ x: 100, y: -200, cx: 1000, cy: 500 }) +
  // The nested group occupies (500, 0)-(1000, 250) of the outer space and maps
  // its own 100 x 100 space onto it: one inner unit is 5 x 2.5 outer units.
  `<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="500" y="0"/><a:ext cx="500" cy="250"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></wpg:grpSpPr>${textBoxShape(
    "Inner",
    { x: 20, y: 40, cx: 60, cy: 40 },
    "Gamma",
  )}</wpg:grpSp>`;

const documentWith = (drawing: string, trailingText = "Anchor"): Document => ({
  package: {
    document: parseDocumentBody(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body><w:p><w:r>${drawing}</w:r><w:r><w:t>${trailingText}</w:t></w:r></w:p></w:body></w:document>`,
    ),
  },
});

const firstParagraph = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected a paragraph");
  }
  return paragraph;
};

const runContents = (paragraph: Paragraph) =>
  paragraph.content.filter((content) => content.type === "run").flatMap((run) => run.content);

const liftedShapes = (paragraph: Paragraph): Shape[] =>
  runContents(paragraph)
    .filter((content): content is ShapeContent => content.type === "shape")
    .map(({ shape }) => shape)
    .filter((shape) => shape.groupChild !== undefined);

const groupDrawing = (paragraph: Paragraph): DrawingContent => {
  const drawing = runContents(paragraph).find((content) => content.type === "drawing");
  if (drawing?.type !== "drawing") {
    throw new Error("Expected the group drawing");
  }
  return drawing;
};

const svgOf = (drawing: DrawingContent): string =>
  decodeURIComponent(drawing.image.src?.split(",").at(1) ?? "");

const textOf = (shape: Shape): string =>
  (shape.textBody?.content ?? [])
    .flatMap((block) => (block.type === "paragraph" ? block.content : []))
    .flatMap((content) => (content.type === "run" ? content.content : []))
    .map((content) => (content.type === "text" ? content.text : ""))
    .join("");

describe("text boxes inside a DrawingML group", () => {
  test("maps each child frame through the group's child coordinate space", () => {
    const paragraph = firstParagraph(documentWith(anchoredGroup(TWO_TEXT_BOXES)));
    const [first, second] = liftedShapes(paragraph);

    // (150 - 100) * 2000 = 100,000 and (-150 + 200) * 2000 = 100,000 from the
    // group's corner, which sits at (500,000, 250,000).
    expect(first?.size).toEqual({ width: 800_000, height: 200_000 });
    expect(first?.position).toEqual({
      horizontal: { relativeTo: "page", posOffset: 600_000 },
      vertical: { relativeTo: "paragraph", posOffset: 350_000 },
    });
    expect(second?.size).toEqual({ width: 600_000, height: 400_000 });
    expect(second?.position).toEqual({
      horizontal: { relativeTo: "page", posOffset: 1_500_000 },
      vertical: { relativeTo: "paragraph", posOffset: 650_000 },
    });
    expect(first && textOf(first)).toBe("Alpha");
    expect(second && textOf(second)).toBe("Beta");
  });

  test("keeps each child's body insets, anchoring and autofit, and the group's layer", () => {
    const paragraph = firstParagraph(documentWith(anchoredGroup(TWO_TEXT_BOXES)));
    const [first, second] = liftedShapes(paragraph);

    expect(first?.textBody?.margins).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    expect(first?.textBody?.anchor).toBe("top");
    expect(second?.textBody?.margins).toEqual({ left: 45_720, top: 0, right: 45_720, bottom: 0 });
    expect(second?.textBody?.anchor).toBe("middle");
    expect(second?.textBody?.autoFit).toBe("none");
    expect(first?.wrap).toEqual({ type: "behind" });
    expect(first?.anchor?.relativeHeight).toBe(7);
    expect(first?.groupChild?.path).toEqual([3]);
    expect(second?.groupChild?.path).toEqual([4]);
  });

  test("leaves the lifted text out of the group's preview but keeps its geometry", () => {
    const drawing = groupDrawing(firstParagraph(documentWith(anchoredGroup(TWO_TEXT_BOXES))));
    const svg = svgOf(drawing);

    expect(svg).toContain('fill="#D9E1F3"');
    expect(svg).not.toContain("Alpha");
    expect(svg).not.toContain("Beta");
  });

  test("composes a nested group's mapping with its parent's", () => {
    const paragraph = firstParagraph(documentWith(anchoredGroup(NESTED)));
    const [inner] = liftedShapes(paragraph);

    // Inner (20, 40) -> outer (500 + 20 * 5, 0 + 40 * 2.5) = (600, 100) ->
    // group EMU ((600 - 100) * 2000, (100 + 200) * 2000) = (1,000,000, 600,000).
    expect(inner?.size).toEqual({ width: 600_000, height: 200_000 });
    expect(inner?.position).toEqual({
      horizontal: { relativeTo: "page", posOffset: 1_500_000 },
      vertical: { relativeTo: "paragraph", posOffset: 850_000 },
    });
    expect(inner?.groupChild?.path).toEqual([3, 2]);
    expect(inner && textOf(inner)).toBe("Gamma");
  });

  test("carries a child's rotation and flips onto its text box", () => {
    const rotated = textBoxShape("Turned", { x: 150, y: -150, cx: 400, cy: 100 }, "Delta").replace(
      "<a:xfrm>",
      '<a:xfrm rot="5400000" flipH="1">',
    );
    const [shape] = liftedShapes(firstParagraph(documentWith(anchoredGroup(rotated))));

    expect(shape?.transform).toEqual({ rotation: 90, flipH: true });
  });

  test("draws a nested group and a zero-height line in the preview's child space", () => {
    const line =
      '<wps:wsp><wps:cNvPr id="7" name="Line"/><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="100" y="-100"/><a:ext cx="1000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="6000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>';
    const nestedFill = NESTED.replace(
      "</wpg:grpSpPr>",
      `</wpg:grpSpPr>${rectangle({ x: 0, y: 0, cx: 100, cy: 100 })}`,
    );
    const svg = svgOf(groupDrawing(firstParagraph(documentWith(anchoredGroup(nestedFill + line)))));

    // 6,000 EMU of line width is 3 units where one unit spans 2,000 EMU.
    expect(svg).toContain(
      '<line x1="100" y1="-100" x2="1100" y2="-100" stroke="#000000" stroke-width="3"/>',
    );
    expect(svg).toContain('<g transform="translate(500 0) scale(5 2.5) translate(0 0)">');
    expect(svg).not.toContain("Gamma");
  });

  test("draws the text of a group placed by alignment in the preview, at its authored size", () => {
    const aligned = anchoredGroup(
      TWO_TEXT_BOXES,
      '<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>250000</wp:posOffset></wp:positionV>',
    );
    const paragraph = firstParagraph(documentWith(aligned));

    expect(liftedShapes(paragraph)).toHaveLength(0);
    const svg = svgOf(groupDrawing(paragraph));
    expect(svg).toContain("Alpha");
    // 11pt is 139,700 EMU, which is 69.85 units of a 2,000-EMU child space.
    expect(svg).toContain(`scale(${139_700 / 2000 / 1000})`);
  });
});

describe("saving text boxes inside a DrawingML group", () => {
  const roundTrip = (document: Document, edit?: (state: EditorState) => EditorState): Paragraph => {
    const pmDocument = toProseDoc(document);
    const state = EditorState.create({ doc: pmDocument });
    const edited = edit ? edit(state).doc : pmDocument;
    return firstParagraph(fromProseDoc(edited, document));
  };

  const textPositionIn = (state: EditorState, name: string): number => {
    let found: number | undefined;
    state.doc.descendants((node, position) => {
      if (node.type.name === "textBox" && node.attrs["textBoxName"] === name) {
        found = position + 2;
      }
    });
    if (found === undefined) {
      throw new Error(`Expected a text box named ${name}`);
    }
    return found;
  };

  test("projects each lifted text box as an editable text box node", () => {
    const pmDocument = toProseDoc(documentWith(anchoredGroup(TWO_TEXT_BOXES)));
    const groupChildren: unknown[] = [];
    pmDocument.descendants((node) => {
      if (node.type.name === "textBox") {
        groupChildren.push(node.attrs["_docxGroupChild"]);
      }
    });

    expect(groupChildren).toHaveLength(2);
    expect(groupChildren.every((child) => child !== null)).toBe(true);
  });

  test("replays the group unchanged and writes no second drawing when nothing was edited", () => {
    const document = documentWith(anchoredGroup(TWO_TEXT_BOXES));
    const sourceXml = groupDrawing(firstParagraph(document)).rawXml ?? "";
    const xml = serializeParagraph(roundTrip(document));

    expect(xml).toContain(sourceXml);
    expect(xml.match(/<wps:txbx>/gu)).toHaveLength(2);
    expect(xml).not.toContain("wordprocessingShape");
  });

  test("writes an edited text box back into its place in the group", () => {
    const document = documentWith(anchoredGroup(NESTED));
    const paragraph = roundTrip(document, (state) =>
      state.apply(state.tr.insertText("Edited ", textPositionIn(state, "Inner"))),
    );
    const xml = serializeParagraph(paragraph);

    expect(xml.match(/<wpg:wgp>/gu)).toHaveLength(1);
    expect(xml.match(/<wps:txbx>/gu)).toHaveLength(1);
    expect(xml).not.toContain("wordprocessingShape");
    const nested = xml.slice(xml.indexOf("<wpg:grpSp>"), xml.indexOf("</wpg:grpSp>"));
    expect(nested).toContain("Edited Gamma");
    expect(xml.indexOf("Edited Gamma")).toBeLessThan(xml.indexOf("</wpg:wgp>"));
  });

  test("keeps the untouched sibling's authored XML when one text box is edited", () => {
    const document = documentWith(anchoredGroup(TWO_TEXT_BOXES));
    const paragraph = roundTrip(document, (state) =>
      state.apply(state.tr.insertText("Edited ", textPositionIn(state, "Second"))),
    );
    const xml = serializeParagraph(paragraph);

    expect(xml).toContain("<w:txbxContent><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:txbxContent>");
    expect(xml).toContain("Edited Beta");
    expect(xml.match(/<wps:txbx>/gu)).toHaveLength(2);
  });

  test("writes an edit into the chosen branch of an alternate-content group", () => {
    const alternate = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wpg">${anchoredGroup(
      TWO_TEXT_BOXES,
    )}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>`;
    const document = documentWith(alternate);
    expect(liftedShapes(firstParagraph(document))).toHaveLength(2);

    const xml = serializeParagraph(
      roundTrip(document, (state) =>
        state.apply(state.tr.insertText("Edited ", textPositionIn(state, "First"))),
      ),
    );

    expect(xml.match(/<mc:AlternateContent/gu)).toHaveLength(1);
    expect(xml).toContain("<mc:Fallback><w:pict/></mc:Fallback>");
    expect(xml.indexOf("Edited Alpha")).toBeGreaterThan(xml.indexOf("<mc:Choice"));
    expect(xml.indexOf("Edited Alpha")).toBeLessThan(xml.indexOf("</mc:Choice>"));
    expect(xml).not.toContain("wordprocessingShape");
  });

  test("writes a text box whose group was deleted as a text box of its own", () => {
    const document = documentWith(anchoredGroup(TWO_TEXT_BOXES));
    const paragraph = roundTrip(document, (state) => {
      let imagePosition: number | undefined;
      state.doc.descendants((node, position) => {
        if (node.type.name === "image") {
          imagePosition = position;
        }
      });
      if (imagePosition === undefined) {
        throw new Error("Expected the group's image node");
      }
      return state.apply(state.tr.delete(imagePosition, imagePosition + 1));
    });
    const xml = serializeParagraph(paragraph);

    expect(xml).not.toContain("<wpg:wgp>");
    expect(xml.match(/wordprocessingShape/gu)).toHaveLength(2);
    expect(xml).toContain("Alpha");
    expect(xml).toContain("Beta");
  });
});
