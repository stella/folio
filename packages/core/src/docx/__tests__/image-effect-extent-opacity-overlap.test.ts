/**
 * Round-trip tests for OOXML image attributes that previously round-tripped
 * incorrectly. Mirrors eigenpal docx-editor PR #424 (sha c605277c9), narrowed
 * to the three sub-items the folio fork still needed:
 *   A. wp:effectExtent vs wp:inline/wp:anchor distT/B/L/R separation
 *   B. a:alphaModFix image opacity (parse + serialize)
 *   C. wp:anchor layoutInCell / allowOverlap as tri-state
 */

import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, Image, Run } from "../../types/document";
import { parseDocumentBody } from "../documentParser";
import { parseDrawing } from "../imageParser";
import { serializeRun } from "../serializer/runSerializer";
import { parseXml } from "../xmlParser";
import type { XmlElement } from "../xmlParser";

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

function parseDrawingFromXml(innerXml: string): Image | null {
  const doc = parseXml(`<w:drawing ${NS}>${innerXml}</w:drawing>`);
  const drawing = (doc.elements as XmlElement[])[0];
  if (!drawing) {
    return null;
  }
  return parseDrawing(drawing, undefined, undefined);
}

function serializeImage(image: Image): string {
  const run: Run = { type: "run", content: [{ type: "drawing", image }] };
  return serializeRun(run);
}

/** Parse the `<w:drawing>` payload out of a serialized `<w:r>...</w:r>` blob. */
function reparseSerializedImage(xml: string): Image | null {
  const wrapped = `<root ${NS}>${xml}</root>`;
  const doc = parseXml(wrapped);
  const root = (doc.elements as XmlElement[])[0];
  if (!root) {
    return null;
  }
  const wr = (root.elements as XmlElement[])[0]; // <w:r>
  if (!wr) {
    return null;
  }
  const drawing = (wr.elements as XmlElement[])[0]; // <w:drawing>
  if (!drawing) {
    return null;
  }
  return parseDrawing(drawing, undefined, undefined);
}

describe("wp:effectExtent stays separate from wp:inline/wp:anchor dist*", () => {
  test("preserves an image drawing whose authored frame has independent geometry", () => {
    const body = parseDocumentBody(`
      <w:document ${NS}><w:body><w:p><w:r><w:drawing>
        <wp:inline><wp:extent cx="1010000" cy="505000"/>
          <wp:effectExtent l="11" t="22" r="33" b="44"/>
          <wp:docPr id="31" name="Frame"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:nvPicPr><pic:cNvPr id="31" name="Frame source"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="rId7" cstate="print"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="7" y="9"/><a:ext cx="1009000" cy="504000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic></a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`);
    const paragraph = body.content.at(0);
    if (paragraph?.type !== "paragraph") {
      throw new Error("Expected parsed paragraph");
    }
    const run = paragraph.content.at(0);
    if (run?.type !== "run") {
      throw new Error("Expected parsed run");
    }
    const drawing = run.content.at(0);
    if (drawing?.type !== "drawing") {
      throw new Error("Expected parsed drawing");
    }

    expect(drawing.rawXml).toBeDefined();
    const document: Document = { package: { document: { content: body.content } } };
    const restored = fromProseDoc(toProseDoc(document), document);
    const restoredParagraph = restored.package.document.content.at(0);
    if (restoredParagraph?.type !== "paragraph") {
      throw new Error("Expected restored paragraph");
    }
    const restoredRun = restoredParagraph.content.at(0);
    if (restoredRun?.type !== "run") {
      throw new Error("Expected restored run");
    }

    const xml = serializeRun(restoredRun);
    expect(xml).toContain('<wp:extent cx="1010000" cy="505000"/>');
    expect(xml).toContain('<wp:effectExtent l="11" t="22" r="33" b="44"/>');
    expect(xml).toContain('<a:off x="7" y="9"/>');
    expect(xml).toContain('<a:ext cx="1009000" cy="504000"/>');
    expect(xml).toContain('cstate="print"');
  });

  test("inline image padding round-trips through <wp:effectExtent>, not dist*", () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 1_000_000, height: 500_000 },
      wrap: { type: "inline" },
      // image.padding is OOXML's wp:effectExtent reservation (EMUs).
      padding: { top: 100, bottom: 200, left: 300, right: 400 },
    });
    expect(xml).toContain('<wp:effectExtent l="300" t="100" r="400" b="200"/>');
    // An omitted distance is semantically zero but remains absent so an
    // untouched parse → save → parse is a model fixed point.
    expect(xml).not.toMatch(/\bdist[TLBR]=/u);
    expect(reparseSerializedImage(xml)?.wrap).toEqual({ type: "inline" });
  });

  test("inline image wrap.dist* serializes to wp:inline dist* attrs", () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline", distT: 1, distB: 2, distL: 3, distR: 4 },
    });
    expect(xml).toContain('distT="1" distB="2" distL="3" distR="4"');
    // No padding set → effectExtent should be all zeros.
    expect(xml).toContain('<wp:effectExtent l="0" t="0" r="0" b="0"/>');
  });

  test("floating image keeps padding and wrap.dist* on independent elements", () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "square", distT: 10, distB: 20, distL: 30, distR: 40 },
      padding: { top: 1, bottom: 2, left: 3, right: 4 },
      position: {
        horizontal: { relativeTo: "column", posOffset: 0 },
        vertical: { relativeTo: "paragraph", posOffset: 0 },
      },
    });
    expect(xml).toContain('distT="10" distB="20" distL="30" distR="40"');
    expect(xml).toContain('<wp:effectExtent l="3" t="1" r="4" b="2"/>');
  });

  test("floating image preserves absent wrap distances", () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "behind" },
      position: {
        horizontal: { relativeTo: "column", posOffset: 0 },
        vertical: { relativeTo: "paragraph", posOffset: 0 },
      },
    });

    expect(xml).toContain("<wp:anchor simplePos=");
    expect(xml).not.toMatch(/\bdist[TLBR]=/u);
    expect(reparseSerializedImage(xml)?.wrap).toEqual({ type: "behind" });
  });

  test("explicit zero wrap distances remain explicit", () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline", distT: 0, distB: 0, distL: 0, distR: 0 },
    });

    expect(xml).toContain('distT="0" distB="0" distL="0" distR="0"');
    expect(reparseSerializedImage(xml)?.wrap).toEqual({
      type: "inline",
      distT: 0,
      distB: 0,
      distL: 0,
      distR: 0,
    });
  });

  test("padding survives a full XML round-trip", () => {
    const original: Image = {
      type: "image",
      rId: "rId1",
      size: { width: 1_000_000, height: 500_000 },
      wrap: { type: "inline" },
      padding: { top: 100, bottom: 200, left: 300, right: 400 },
    };
    const xml = serializeImage(original);
    const parsed = reparseSerializedImage(xml);
    expect(parsed?.padding).toEqual(original.padding);
  });
});

describe("a:alphaModFix opacity round-trip", () => {
  test("parse a:alphaModFix amt as opacity fraction", () => {
    const img = parseDrawingFromXml(`
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="1000000" cy="500000"/>
        <wp:docPr id="1" name="Picture 1"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="rId1"><a:alphaModFix amt="50000"/></a:blip>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>`);
    expect(img?.opacity).toBeCloseTo(0.5, 5);
  });

  test('amt="100000" (fully opaque) does not produce an opacity field', () => {
    const img = parseDrawingFromXml(`
      <wp:inline>
        <wp:extent cx="100" cy="100"/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"><a:alphaModFix amt="100000"/></a:blip></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:inline>`);
    expect(img?.opacity).toBeUndefined();
  });

  test('non-numeric amt (e.g. amt="oops") parses as undefined, not NaN', () => {
    const img = parseDrawingFromXml(`
      <wp:inline>
        <wp:extent cx="100" cy="100"/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"><a:alphaModFix amt="oops"/></a:blip></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:inline>`);
    expect(img?.opacity).toBeUndefined();
  });

  test("serialize opacity < 1 emits a:alphaModFix; opacity 1 omits it", () => {
    const opaque = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline" },
    });
    expect(opaque).not.toContain("alphaModFix");

    const transparent = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline" },
      opacity: 0.5,
    });
    expect(transparent).toContain('<a:alphaModFix amt="50000"/>');
  });

  test("opacity round-trips through XML", () => {
    const original: Image = {
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline" },
      opacity: 0.25,
    };
    const xml = serializeImage(original);
    const parsed = reparseSerializedImage(xml);
    expect(parsed?.opacity).toBeCloseTo(0.25, 5);
  });
});

describe("a:lum brightness and contrast round-trip", () => {
  test("parses signed percentages by namespace URI", () => {
    const img = parseDrawingFromXml(`
      <wp:inline>
        <wp:extent cx="100" cy="100"/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1">
              <effects:lum xmlns:effects="http://purl.oclc.org/ooxml/drawingml/main"
                bright="70.001%" contrast="-70%"/>
            </a:blip></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:inline>`);

    expect(img?.effects).toEqual({ brightness: 70.001, contrast: -70 });
  });

  test("survives the editable model and regenerated DrawingML", () => {
    const image: Image = {
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline" },
      effects: { brightness: 70.001, contrast: -70 },
    };
    const document: Document = {
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
    };

    const restored = fromProseDoc(toProseDoc(document), document);
    const paragraph = restored.package.document.content.at(0);
    const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
    const drawing = run?.type === "run" ? run.content.at(0) : undefined;
    if (drawing?.type !== "drawing" || !drawing.image) {
      throw new Error("Expected restored image drawing");
    }

    expect(drawing.image.effects).toEqual(image.effects);
    expect(serializeImage(drawing.image)).toContain('<a:lum bright="70001" contrast="-70000"/>');
  });

  test("preserves explicit zeroes when regenerating", () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "inline" },
      effects: { brightness: 0, contrast: 0 },
    });
    expect(xml).toContain('<a:lum bright="0" contrast="0"/>');
  });
});

describe("wp:anchor layoutInCell / allowOverlap tri-state round-trip", () => {
  test('parse explicit "0" → false', () => {
    const img = parseDrawingFromXml(`
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                 relativeHeight="0" behindDoc="0" locked="0"
                 layoutInCell="0" allowOverlap="0">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="100" cy="100"/>
        <wp:wrapNone/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:anchor>`);
    expect(img?.anchor?.layoutInCell).toBe(false);
    expect(img?.anchor?.allowOverlap).toBe(false);
  });

  test('parse "true"/"false" literals (OOXML ST_OnOff full set)', () => {
    const trueImg = parseDrawingFromXml(`
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                 relativeHeight="0" behindDoc="0" locked="0"
                 layoutInCell="true" allowOverlap="true">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="100" cy="100"/>
        <wp:wrapNone/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:anchor>`);
    expect(trueImg?.anchor?.layoutInCell).toBe(true);
    expect(trueImg?.anchor?.allowOverlap).toBe(true);

    const falseImg = parseDrawingFromXml(`
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                 relativeHeight="0" behindDoc="0" locked="0"
                 layoutInCell="false" allowOverlap="false">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="100" cy="100"/>
        <wp:wrapNone/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:anchor>`);
    expect(falseImg?.anchor?.layoutInCell).toBe(false);
    expect(falseImg?.anchor?.allowOverlap).toBe(false);
  });

  test("parse absent attrs → undefined (omit the field entirely)", () => {
    const img = parseDrawingFromXml(`
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                 relativeHeight="0" behindDoc="0" locked="0">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="100" cy="100"/>
        <wp:wrapNone/>
        <wp:docPr id="1" name="img"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:ext cx="100" cy="100"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData></a:graphic>
      </wp:anchor>`);
    expect(img?.anchor?.layoutInCell).toBeUndefined();
    expect(img?.anchor?.allowOverlap).toBeUndefined();
  });

  test('serializer emits explicit "0" only when the model says false', () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "square" },
      position: {
        horizontal: { relativeTo: "column", posOffset: 0 },
        vertical: { relativeTo: "paragraph", posOffset: 0 },
      },
      anchor: { layoutInCell: false, allowOverlap: false },
    });
    expect(xml).toContain('layoutInCell="0"');
    expect(xml).toContain('allowOverlap="0"');
  });

  test('absent or explicit-true folds back to the spec default "1"', () => {
    const xml = serializeImage({
      type: "image",
      rId: "rId1",
      size: { width: 100, height: 100 },
      wrap: { type: "square" },
      position: {
        horizontal: { relativeTo: "column", posOffset: 0 },
        vertical: { relativeTo: "paragraph", posOffset: 0 },
      },
    });
    expect(xml).toContain('layoutInCell="1"');
    expect(xml).toContain('allowOverlap="1"');
  });
});
