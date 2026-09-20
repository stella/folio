/**
 * An edit inside a bidirectional wrapper goes out through the selective patcher.
 *
 * Selective save splices only the paragraphs the editor changed, so the
 * paragraph that held the wrapper is rebuilt from the model while its
 * neighbours keep their original bytes. That makes it the path a real edit
 * takes, and the one where a wrapper the save leg forgot would disappear from
 * a file the user still sees as unchanged everywhere else.
 *
 * The comment range inside the wrapper is here for the splice guard: a patch
 * that emitted one half of a range would be refused
 * (`patchBreaksCommentRangeBalance`), and the wrapper must not move a marker
 * out of the paragraph it belongs to.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState, TextSelection } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { attemptSelectiveSave } from "./selectiveSave";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const CONTENT_TYPES = `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const PACKAGE_RELS = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const COMMENTS = `${XML_DECLARATION}
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="R">
    <w:p><w:r><w:t>A note</w:t></w:r></w:p>
  </w:comment>
</w:comments>`;

const CORE_PROPERTIES = `${XML_DECLARATION}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Inline wrapper fixture</dc:title>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified>
</cp:coreProperties>`;

const WRAPPED_PARAGRAPH_ID = "10000001";
const UNTOUCHED_PARAGRAPH_ID = "10000002";

const DOCUMENT = `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="${WRAPPED_PARAGRAPH_ID}"><w:r><w:t>before </w:t></w:r><w:bdo w:val="rtl"><w:commentRangeStart w:id="1"/><w:r><w:t>inside</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:bdo><w:r><w:t> after</w:t></w:r></w:p>
    <w:p w14:paraId="${UNTOUCHED_PARAGRAPH_ID}"><w:r><w:t>untouched</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;

const fixture = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", PACKAGE_RELS);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/document.xml", DOCUMENT);
  zip.file("word/comments.xml", COMMENTS);
  zip.file("docProps/core.xml", CORE_PROPERTIES);
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXmlOf = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const text = await zip.file("word/document.xml")?.async("text");
  if (text === undefined) {
    throw new Error("The saved package has no word/document.xml");
  }
  return text;
};

const rangeOfText = (state: EditorState, text: string): { from: number; to: number } => {
  let range: { from: number; to: number } | undefined;
  state.doc.descendants((node, position) => {
    const at = node.isText ? (node.text?.indexOf(text) ?? -1) : -1;
    if (range === undefined && at >= 0) {
      range = { from: position + at, to: position + at + text.length };
    }
  });
  if (range === undefined) {
    throw new Error(`The editor holds no text ${text}`);
  }
  return range;
};

describe("a selective save of an edit inside a bidirectional wrapper", () => {
  test("patches the paragraph with the wrapper still around the edited text", async () => {
    const buffer = await fixture();
    const parsed = await parseDocx(buffer, { preloadFonts: false });
    const state = EditorState.create({ doc: toProseDoc(parsed) });
    const { from, to } = rangeOfText(state, "side");
    const edited = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to)).insertText("SIDE", from, to),
    );

    const saved = await attemptSelectiveSave(fromProseDoc(edited.doc, parsed), buffer, {
      changedParaIds: new Set([WRAPPED_PARAGRAPH_ID]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(saved).not.toBeNull();
    if (!saved) {
      throw new Error("The selective save refused the patch");
    }

    const xml = await documentXmlOf(saved);
    expect(xml).toContain('<w:bdo w:val="rtl">');
    expect(xml).toContain("SIDE");
    // The splice guard refuses a patch that leaves one half of a range behind.
    expect(xml.match(/<w:commentRangeStart\b/gu)).toHaveLength(1);
    expect(xml.match(/<w:commentRangeEnd\b/gu)).toHaveLength(1);
    // Only the edited paragraph was rebuilt.
    expect(xml).toContain(`<w:p w14:paraId="${UNTOUCHED_PARAGRAPH_ID}"><w:r><w:t>untouched</w:t>`);
  });
});
