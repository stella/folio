import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { Fragment, type Node as PMNode } from "prosemirror-model";

import { mergeImageAttrs } from "../prosemirror/attrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx, repackDocx } from "../docx/rezip";
import { prepareTargetInlineAtom } from "./inline-atom-resources";

const IMAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const pngBytes = (base64: string): Uint8Array =>
  Uint8Array.from(atob(base64), (character) => character.codePointAt(0) ?? 0);
const TARGET_IMAGE_BYTES = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
);
const BASE_IMAGE_BYTES = pngBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
);

const documentXml = (body: string): string => `
  <w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <w:body>${body}<w:sectPr/></w:body>
  </w:document>`;

const targetDrawing = `<w:drawing><wp:inline>
  <wp:extent cx="1001000" cy="500500"/><wp:effectExtent l="5" t="7" r="11" b="13"/>
  <wp:docPr id="9" name="Target picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:pic><pic:nvPicPr><pic:cNvPr id="9" name="Target source"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="rIdTarget" cstate="print"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="17" y="19"/><a:ext cx="999000" cy="499000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic>
  </wp:inline></w:drawing>`;

const ordinaryDrawing = `<w:drawing><wp:inline>
  <wp:extent cx="251460" cy="243840"/><wp:docPr id="12" name="Ordinary picture"/>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:pic><pic:nvPicPr><pic:cNvPr id="12" name="ordinary.png"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="rIdOrdinary"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="251460" cy="243840"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic>
  </wp:inline></w:drawing>`;

const packageWithDocument = async ({
  body,
  relationshipId,
  mediaName,
  media,
}: {
  body: string;
  relationshipId: string;
  mediaName: string;
  media: Uint8Array;
}): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentXml(body));
  const contentTypesPath = "[Content_Types].xml";
  const contentTypes = await zip.file(contentTypesPath)!.async("text");
  zip.file(
    contentTypesPath,
    contentTypes.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'),
  );
  const relsPath = "word/_rels/document.xml.rels";
  const rels = await zip.file(relsPath)!.async("text");
  zip.file(
    relsPath,
    rels.replace(
      "</Relationships>",
      `<Relationship Id="${relationshipId}" Type="${IMAGE_RELATIONSHIP}" Target="media/${mediaName}"/></Relationships>`,
    ),
  );
  zip.file(`word/media/${mediaName}`, media);
  return zip.generateAsync({ type: "arraybuffer" });
};

const firstImageNode = (document: PMNode): PMNode => {
  let image: PMNode | undefined;
  document.descendants((node) => {
    if (node.type.name === "image") {
      image = node;
      return false;
    }
    return true;
  });
  if (!image) {
    throw new Error("Expected parsed target image");
  }
  return image;
};

describe("prepareTargetInlineAtom image resources", () => {
  test("imports a parsed raw target picture without changing its drawing geometry", async () => {
    const target = await parseDocx(
      await packageWithDocument({
        body: `<w:p><w:r>${targetDrawing}</w:r></w:p>`,
        relationshipId: "rIdTarget",
        mediaName: "target.png",
        media: TARGET_IMAGE_BYTES,
      }),
      { preloadFonts: false },
    );
    const base = await parseDocx(
      await packageWithDocument({
        body: `<w:p><w:r><w:t>Base</w:t></w:r><w:r>${targetDrawing.replaceAll("rIdTarget", "rId1")}</w:r></w:p>`,
        relationshipId: "rId1",
        mediaName: "image1.png",
        media: BASE_IMAGE_BYTES,
      }),
      { preloadFonts: false },
    );

    expect(target.package.relationships?.get("rIdTarget")?.target).toBe("media/target.png");
    expect([...target.package.media!.keys()]).toContain("word/media/target.png");
    expect(target.package.media?.get("word/media/target.png")?.dataUrl).toStartWith("data:");
    const targetImage = firstImageNode(toProseDoc(target));
    expect(targetImage.attrs.src).toStartWith("data:");
    expect(targetImage.attrs._docxRawXml).toContain('r:embed="rIdTarget"');
    const prepared = prepareTargetInlineAtom(targetImage);
    if (!prepared) {
      throw new Error("Expected supported target image atom");
    }

    const basePm = toProseDoc(base);
    const paragraph = basePm.firstChild;
    if (!paragraph) {
      throw new Error("Expected base paragraph");
    }
    const mergedPm = basePm.copy(
      basePm.content.replaceChild(
        0,
        paragraph.copy(paragraph.content.append(Fragment.from(prepared))),
      ),
    );
    const saved = await repackDocx(fromProseDoc(mergedPm, base), { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file("word/document.xml")!.async("text");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("text");

    expect(xml).toContain('<wp:extent cx="1001000" cy="500500"/>');
    expect(xml).toContain('<wp:effectExtent l="5" t="7" r="11" b="13"/>');
    expect(xml).toContain('<a:off x="17" y="19"/>');
    expect(xml).toContain('<a:ext cx="999000" cy="499000"/>');
    expect(xml).toContain('cstate="print"');
    expect(xml).not.toContain('r:embed="rIdTarget"');
    const docPrIds = [...xml.matchAll(/<wp:docPr id="(?<id>\d+)"/gu)].map(
      (match) => match.groups?.id,
    );
    expect(docPrIds).toContain("9");
    expect(docPrIds).toContain("100000");
    expect(new Set(docPrIds).size).toBe(docPrIds.length);
    expect(xml).toContain('<pic:cNvPr id="100000" name="Target source"/>');
    expect(rels).toContain('Id="rId1"');
    const imageRelationships = [
      ...rels.matchAll(
        /<Relationship Id="(?<id>rId\d+)" Type="[^"]*\/image" Target="media\/(?<file>[^"]+)"\/>/gu,
      ),
    ];
    const imported = imageRelationships.find((relationship) => relationship.groups?.id !== "rId1");
    expect(imported?.groups?.id).toBeDefined();
    expect(imported?.groups?.file).toBeDefined();
    const importedBytes = await zip
      .file(`word/media/${imported?.groups?.file}`)!
      .async("uint8array");
    expect(importedBytes).toEqual(TARGET_IMAGE_BYTES);
    expect(await zip.file("word/media/image1.png")!.async("uint8array")).toEqual(BASE_IMAGE_BYTES);
  });

  test("keeps ordinary parsed picture EMUs byte-for-byte when unchanged", async () => {
    const parsed = await parseDocx(
      await packageWithDocument({
        body: `<w:p><w:r>${ordinaryDrawing}</w:r></w:p>`,
        relationshipId: "rIdOrdinary",
        mediaName: "ordinary.png",
        media: TARGET_IMAGE_BYTES,
      }),
      { preloadFonts: false },
    );

    const paragraph = parsed.package.document.content.find((block) => block.type === "paragraph");
    if (!paragraph) {
      throw new Error("Expected parsed paragraph");
    }
    const run = paragraph.content.find((content) => content.type === "run");
    if (!run) {
      throw new Error("Expected parsed run");
    }
    const drawing = run.content.find((content) => content.type === "drawing");
    if (!drawing) {
      throw new Error("Expected parsed drawing");
    }
    expect(drawing.rawXml).toContain('cx="251460"');
    expect(drawing.rawImageFingerprint).toBeDefined();

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file("word/document.xml")!.async("text");
    expect(xml).toContain('<wp:extent cx="251460" cy="243840"/>');
    expect(xml).toContain('<a:ext cx="251460" cy="243840"/>');
  });

  test("serializes direct modeled image edits instead of stale raw drawing XML", async () => {
    const parsed = await parseDocx(
      await packageWithDocument({
        body: `<w:p><w:r>${ordinaryDrawing}</w:r></w:p>`,
        relationshipId: "rIdOrdinary",
        mediaName: "ordinary.png",
        media: TARGET_IMAGE_BYTES,
      }),
      { preloadFonts: false },
    );
    const paragraph = parsed.package.document.content.find((block) => block.type === "paragraph");
    if (!paragraph) {
      throw new Error("Expected parsed paragraph");
    }
    const run = paragraph.content.find((content) => content.type === "run");
    if (!run) {
      throw new Error("Expected parsed run");
    }
    const drawing = run.content.find((content) => content.type === "drawing");
    if (!drawing) {
      throw new Error("Expected parsed drawing");
    }
    drawing.image.size = { width: 777000, height: 333000 };

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file("word/document.xml")!.async("text");
    expect(xml).toContain('<wp:extent cx="777000" cy="333000"/>');
    expect(xml).toContain('<a:ext cx="777000" cy="333000"/>');
    expect(xml).not.toContain('<wp:extent cx="251460" cy="243840"/>');
  });

  test("serializes ProseMirror image edits instead of stale raw drawing XML", async () => {
    const parsed = await parseDocx(
      await packageWithDocument({
        body: `<w:p><w:r>${ordinaryDrawing}</w:r></w:p>`,
        relationshipId: "rIdOrdinary",
        mediaName: "ordinary.png",
        media: TARGET_IMAGE_BYTES,
      }),
      { preloadFonts: false },
    );
    const pm = toProseDoc(parsed);
    const paragraph = pm.firstChild;
    if (!paragraph) {
      throw new Error("Expected parsed ProseMirror paragraph");
    }
    const image = firstImageNode(pm);
    const editedImage = image.type.create(mergeImageAttrs(image, { width: 100, height: 50 }));
    const edited = pm.copy(pm.content.replaceChild(0, paragraph.copy(Fragment.from(editedImage))));

    const saved = await repackDocx(fromProseDoc(edited, parsed), { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file("word/document.xml")!.async("text");
    expect(xml).toContain('<wp:extent cx="952500" cy="476250"/>');
    expect(xml).not.toContain('<wp:extent cx="251460" cy="243840"/>');
  });
});
