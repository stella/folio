import { describe, expect, setDefaultTimeout, test } from "bun:test";
import JSZip from "jszip";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createEmptyDocument } from "../utils/createDocument";

import type { Document, DrawingContent } from "../types/document";
import { normalizeDrawingIds } from "./drawingIdNormalization";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(10_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const textBoxDrawing = (text: string, id?: string): string => `
  <w:drawing>
    <wp:inline>
      <wp:extent cx="914400" cy="457200"/>
      ${id === undefined ? "" : `<wp:docPr id="${id}" name="Text box ${id}"/>`}
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;

const buildDocx = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
    <w:document
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body>
        <w:p><w:r>${textBoxDrawing("Generated")}</w:r></w:p>
        <w:p><w:r>${textBoxDrawing("Authored", "100000")}</w:r></w:p>
        <w:p><w:r>${textBoxDrawing("Zero", "0")}</w:r></w:p>
        <w:sectPr/>
      </w:body>
    </w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const shapeIds = ({ package: { document } }: Document): (string | undefined)[] =>
  document.content.flatMap((block) => {
    if (block.type !== "paragraph") {
      return [];
    }
    return block.content.flatMap((content) => {
      if (content.type !== "run") {
        return [];
      }
      return content.content.flatMap((runContent) =>
        runContent.type === "shape" ? [runContent.shape.id] : [],
      );
    });
  });

describe("drawing ID normalization", () => {
  test("assigns missing shape IDs without colliding and remains stable after save", async () => {
    const parsed = await parseDocx(await buildDocx(), { preloadFonts: false });
    expect(shapeIds(parsed)).toEqual(["100001", "100000", "0"]);

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(shapeIds(reopened)).toEqual(["100001", "100000", "0"]);
  });

  test("reassigns a detached header drawing that collides with a main-story drawing", () => {
    const rawDrawing = ({ id, rId }: { id: string; rId: string }): DrawingContent => ({
      type: "drawing",
      image: {
        type: "image",
        id,
        rId,
        src: "data:image/png;base64,AA==",
        size: { width: 9_144, height: 4_572 },
        wrap: { type: "inline" },
      },
      rawXml: `<w:drawing><wp:inline><wp:docPr id="${id}" name="Picture"/><pic:cNvPr id="${id}" name="Picture"/><a:blip r:embed="${rId}"/></wp:inline></w:drawing>`,
    });
    const main = rawDrawing({ id: "9", rId: "rIdMain" });
    const importedHeader = rawDrawing({ id: "9", rId: "rId_img_compare" });
    const documentBody = {
      content: [
        { type: "paragraph" as const, content: [{ type: "run" as const, content: [main] }] },
      ],
    };
    const headers = new Map([
      [
        "header1.xml",
        {
          type: "header" as const,
          hdrFtrType: "default" as const,
          content: [
            {
              type: "paragraph" as const,
              content: [{ type: "run" as const, content: [importedHeader] }],
            },
          ],
        },
      ],
    ]);

    normalizeDrawingIds({ documentBody, headers });

    expect(main.rawXml).toContain('wp:docPr id="9"');
    expect(importedHeader.rawXml).toContain('wp:docPr id="100000"');
    expect(importedHeader.rawXml).toContain('pic:cNvPr id="100000"');
  });
});

test("converted lexical drawing ids avoid every authored numeric id and reach a fixed point", () => {
  assertProperty(
    fc.property(
      fc.array(fc.stringMatching(/^[A-Za-z_][A-Za-z_0-9]{0,20}$/u), {
        minLength: 1,
        maxLength: 20,
      }),
      fc.constantFrom("", "+", "000"),
      (ids, prefix) => {
        const document = createEmptyDocument();
        const authored = [
          `${prefix}0`,
          `${prefix}4294967295`,
          `${prefix}100000`,
          `${prefix}100001`,
        ];
        const drawings = [...authored, ...ids].map((id) => ({
          type: "drawing" as const,
          image: { rId: "rId1", id, size: { width: 914_400, height: 914_400 } },
        }));
        document.package.document.content = [
          { type: "paragraph", content: [{ type: "run", content: drawings }] },
        ];
        const surfaces = { documentBody: document.package.document };
        normalizeDrawingIds(surfaces);
        const first = drawings.map(({ image }) => image.id);
        expect(first.slice(0, authored.length)).toEqual(authored);
        expect(new Set(first.map(Number)).size).toBe(first.length);
        for (const id of first) {
          expect(Number(id)).toBeGreaterThanOrEqual(0);
          expect(Number(id)).toBeLessThanOrEqual(0xffff_ffff);
        }
        normalizeDrawingIds(surfaces);
        expect(drawings.map(({ image }) => image.id)).toEqual(first);
      },
    ),
    { numRuns: 40 },
  );
});
