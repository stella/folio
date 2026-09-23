/**
 * A picture in a footnote or endnote resolves its `r:embed` against the note
 * part's own relationships and carries its media data, whether it sits in a
 * note paragraph, a table cell, or a block SDT; saving keeps the media part
 * and the note relationship.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { BlockContent, Endnote, Footnote, Image } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NOTE_NAMESPACES = `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

// 1x1 transparent PNG.
const PNG_BYTES = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
  ),
  (char) => char.charCodeAt(0),
);
const MEDIA_PATH = "word/media/note.png";
// Note parts number their relationships independently of the document; the
// document part deliberately has no relationship with this id.
const NOTE_IMAGE_RID = "rId7";

const picture = (docPrId: number): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="${docPrId}" name="Picture ${docPrId}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="Picture ${docPrId}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${NOTE_IMAGE_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

// One note holding a picture in a paragraph, in a table cell, and in a
// paragraph inside a block SDT.
const noteBody = (idBase: number): string =>
  `<w:p>${picture(idBase + 1)}</w:p>` +
  `<w:tbl><w:tr><w:tc><w:p>${picture(idBase + 2)}</w:p></w:tc></w:tr></w:tbl>` +
  `<w:sdt><w:sdtPr/><w:sdtContent><w:p>${picture(idBase + 3)}</w:p></w:sdtContent></w:sdt>`;

const separators = (kind: "footnote" | "endnote"): string =>
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

const notesXml = (kind: "footnote" | "endnote", idBase: number): string =>
  `${XML_DECLARATION}<w:${kind}s ${NOTE_NAMESPACES}>${separators(kind)}<w:${kind} w:id="1">${noteBody(idBase)}</w:${kind}></w:${kind}s>`;

const noteRelsXml = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${NOTE_IMAGE_RID}" Type="${RELATIONSHIP_TYPES.image}" Target="media/note.png"/></Relationships>`;

const documentXml = `${XML_DECLARATION}<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`;

const contentTypesXml = `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>`;

const packageRelsXml = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`;

const documentRelsXml = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/><Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.endnotes}" Target="endnotes.xml"/></Relationships>`;

const createPackage = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", packageRelsXml);
  zip.file("word/_rels/document.xml.rels", documentRelsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/footnotes.xml", notesXml("footnote", 100));
  zip.file("word/_rels/footnotes.xml.rels", noteRelsXml);
  zip.file("word/endnotes.xml", notesXml("endnote", 200));
  zip.file("word/_rels/endnotes.xml.rels", noteRelsXml);
  zip.file(MEDIA_PATH, PNG_BYTES);
  return zip.generateAsync({ type: "arraybuffer" });
};

type NoteImage = { location: "paragraph" | "table" | "sdt"; image: Image };

const paragraphImages = (block: BlockContent): Image[] =>
  block.type === "paragraph"
    ? block.content
        .flatMap((item) => (item.type === "run" ? item.content : []))
        .flatMap((content) => (content.type === "drawing" ? [content.image] : []))
    : [];

const noteImages = (note: Footnote | Endnote | undefined): NoteImage[] => {
  const images: NoteImage[] = [];
  const collect = (location: NoteImage["location"], blocks: BlockContent[]): void => {
    for (const image of blocks.flatMap(paragraphImages)) {
      images.push({ location, image });
    }
  };
  for (const block of note?.content ?? []) {
    switch (block.type) {
      case "paragraph":
        collect("paragraph", [block]);
        break;
      case "table":
        collect(
          "table",
          block.rows.flatMap((row) => row.cells).flatMap((cell) => cell.content),
        );
        break;
      case "blockSdt":
        collect("sdt", block.content);
        break;
    }
  }
  return images;
};

const expectResolvedPictures = (note: Footnote | Endnote | undefined): void => {
  const images = noteImages(note);
  expect(images.map(({ location }) => location)).toEqual(["paragraph", "table", "sdt"]);
  for (const { image } of images) {
    expect(image.rId).toBe(NOTE_IMAGE_RID);
    expect(image.mimeType).toBe("image/png");
    expect(image.src?.startsWith("data:image/png")).toBe(true);
  }
};

describe("pictures in note parts", () => {
  test("footnote and endnote pictures carry media data in every block container", async () => {
    const doc = await parseDocx(await createPackage(), { preloadFonts: false });

    expectResolvedPictures(doc.package.footnotes?.at(0));
    expectResolvedPictures(doc.package.endnotes?.at(0));
  });

  test("saving keeps the note media part and relationships", async () => {
    const doc = await parseDocx(await createPackage(), { preloadFonts: false });
    const zip = await JSZip.loadAsync(await createDocx(doc));

    expect(await zip.file(MEDIA_PATH)?.async("uint8array")).toEqual(PNG_BYTES);
    for (const relsPath of ["word/_rels/footnotes.xml.rels", "word/_rels/endnotes.xml.rels"]) {
      expect(await zip.file(relsPath)?.async("text")).toContain(`Id="${NOTE_IMAGE_RID}"`);
    }

    const reparsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    expectResolvedPictures(reparsed.package.footnotes?.at(0));
    expectResolvedPictures(reparsed.package.endnotes?.at(0));
  });
});
