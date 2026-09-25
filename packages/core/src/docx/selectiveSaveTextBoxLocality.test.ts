/**
 * An edit beside a text box stays one paragraph's edit.
 *
 * A text box read from `mc:AlternateContent` is written twice by its producer:
 * the DrawingML `mc:Choice` the model reads and a VML `mc:Fallback` copy the
 * model does not own. The model's serialization of the part therefore holds a
 * different set of `<w:p>` elements than the source (the Fallback's are
 * absent), and a patch that demanded the two agree paragraph for paragraph
 * refused every edit to such a document, handing it to the full repack, which
 * rebuilds every paragraph and drops each Fallback. The selective patch
 * routes each changed paragraph by its own identity instead, so the rest of
 * the part, text boxes and Fallbacks included, keeps its source bytes.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const CONTENT_TYPES = `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const ROOT_RELS = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const DOCUMENT_RELS = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

const ROOT_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'mc:Ignorable="w14"',
].join(" ");

const MARKER = "‸";

type Coverage = "authored" | "none";

const paragraph = (coverage: Coverage, id: string, content: string): string =>
  `<w:p${coverage === "authored" ? ` w14:paraId="${id}" w14:textId="${id}"` : ""}>${content}</w:p>`;

const textRun = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;

/**
 * A DrawingML text box with its VML Fallback. The Fallback paragraph repeats
 * the Choice paragraph's id, as a producer that regenerates the Fallback does.
 */
const textBoxRun = (coverage: Coverage, text: string): string => {
  const boxParagraph = paragraph(coverage, "10000003", textRun(text));
  return (
    "<w:r><mc:AlternateContent>" +
    '<mc:Choice Requires="wps"><w:drawing><wp:inline>' +
    '<wp:extent cx="914400" cy="457200"/><wp:docPr id="11" name="Text Box 11"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
    `<wps:txbx><w:txbxContent>${boxParagraph}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
    "</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
    '<mc:Fallback><w:pict><v:shape id="Text Box 11" style="width:72pt;height:36pt">' +
    `<v:textbox><w:txbxContent>${boxParagraph}</w:txbxContent></v:textbox>` +
    "</v:shape></w:pict></mc:Fallback>" +
    "</mc:AlternateContent></w:r>"
  );
};

const documentXml = (coverage: Coverage): string =>
  `${XML_DECLARATION}<w:document ${ROOT_NAMESPACES}><w:body>` +
  paragraph(coverage, "10000001", textRun("Before the box")) +
  paragraph(coverage, "10000002", textRun("Host ") + textBoxRun(coverage, "Box text")) +
  paragraph(coverage, "10000004", textRun("After the box")) +
  "<w:sectPr/></w:body></w:document>";

const buildPackage = async (xml: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/document.xml", xml);
  const bytes = await zip.generateAsync({ type: "uint8array" });
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const documentPartOf = async (docx: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(docx)).file("word/document.xml");
  if (!entry) {
    throw new Error("saved package has no word/document.xml");
  }
  return entry.async("text");
};

describe("a body edit beside a text box with a Fallback", () => {
  for (const coverage of ["authored", "none"] as const) {
    test(`keeps every other paragraph byte-exact (paraIds: ${coverage})`, async () => {
      const source = documentXml(coverage);
      const reviewer = await FolioDocxReviewer.fromBuffer(await buildPackage(source));
      const target = reviewer.snapshot().blocks[0];
      if (!target) {
        throw new Error("fixture has no first block");
      }
      const applied = reviewer.applyOperations(
        [
          {
            id: "edit",
            type: "replaceInBlock",
            blockId: target.id,
            find: target.text,
            replace: `${target.text}${MARKER}`,
          },
        ],
        { mode: "direct" },
      );
      expect(applied.applied.length).toBe(1);

      const saved = await documentPartOf(await reviewer.toBuffer());

      // Everything from the host paragraph on — the text box, its Fallback
      // and the trailing paragraph — is the source's own bytes.
      const hostStart = source.indexOf("<w:p", source.indexOf("</w:p>"));
      const tail = source.slice(hostStart);
      expect(saved.endsWith(tail)).toBe(true);
      // So is everything before the edited paragraph.
      const head = source.slice(0, source.indexOf("<w:p"));
      expect(saved.startsWith(head)).toBe(true);
      // And the edited paragraph carries the edit, with no id the source lacked.
      const edited = saved.slice(head.length, saved.length - tail.length);
      expect(edited).toContain(`Before the box${MARKER}`);
      expect(edited.includes("w14:paraId")).toBe(coverage === "authored");
    });
  }
});
