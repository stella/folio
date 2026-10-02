import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  getSourceReplayToken,
  INHERIT_RUN_PROPS,
  normalizeForOps,
  OP_STORIES,
  paragraphLogicalText,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

const WORD_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_2010_URI = "http://schemas.microsoft.com/office/word/2010/wordml";
const RELATIONSHIP_URI = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const UNTOUCHED_BLOCK =
  '<w:p w14:paraId="00000001" custom="keep"><w:r><w:t>First</w:t></w:r></w:p>';
const CHANGED_BLOCK = '<w:p w14:paraId="00000002"><w:r><w:t>Second</w:t></w:r></w:p>';
const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${WORD_URI}" xmlns:w14="${WORD_2010_URI}" xmlns:e="urn:extension" e:flag="keep"><e:before/><w:body>\n${UNTOUCHED_BLOCK}\n${CHANGED_BLOCK}\n</w:body><e:after/></w:document>`;

const fixtureBuffer = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="${RELATIONSHIP_URI}"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file("word/document.xml", DOCUMENT_XML);
  zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="${RELATIONSHIP_URI}"/>`);
  return zip.generateAsync({ type: "arraybuffer" });
};

const savedXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const part = zip.file("word/document.xml");
  if (!part) panic("Expected saved document.xml");
  return part.async("text");
};

const repackTracked = (document: Document): Promise<ArrayBuffer> => {
  const sourceReplay = getSourceReplayToken(document);
  if (sourceReplay === undefined) panic("Expected tracked source replay identity");
  return repackDocx(document, { sourceReplay, updateModifiedDate: false });
};

const secondParagraph = (document: Document) => {
  const paragraph = document.package.document.content.at(1);
  if (paragraph?.type !== "paragraph") panic("Expected second paragraph");
  return paragraph;
};

describe("tracked editor DOCX saves", () => {
  test("an unchanged editor projection preserves document.xml byte for byte", async () => {
    const original = await parseDocx(await fixtureBuffer(), { sourceReplay: "tracked" });
    const token = getSourceReplayToken(original);
    expect(token).toBeDefined();
    const current = fromProseDoc(toProseDoc(original), original);
    expect(getSourceReplayToken(current)).toBe(token);
    expect(current.package.document.content.at(0)).toBe(original.package.document.content.at(0));
    expect(await savedXml(await repackTracked(current))).toBe(DOCUMENT_XML);
  });

  test("one editor edit retains untouched block bytes and leaves the base graph intact", async () => {
    const original = await parseDocx(await fixtureBuffer(), { sourceReplay: "tracked" });
    const baseContent = structuredClone(original.package.document.content);
    const state = EditorState.create({ doc: toProseDoc(original) });
    const edited = fromProseDoc(
      state.tr.insertText("!", state.doc.child(0).nodeSize + 1).doc,
      original,
    );
    expect(edited.package.document.content.at(0)).toBe(original.package.document.content.at(0));
    expect(secondParagraph(edited)).not.toBe(secondParagraph(original));
    const xml = await savedXml(await repackTracked(edited));
    expect(xml).toContain(UNTOUCHED_BLOCK);
    expect(xml).toContain("!Second");
    expect(xml).toContain("<e:before/>");
    expect(xml).toContain("<e:after/>");
    expect(original.package.document.content).toEqual(baseContent);
    expect(paragraphLogicalText(secondParagraph(original))).toBe("Second");
    expect(await savedXml(await repackTracked(original))).toBe(DOCUMENT_XML);
  });

  test("an untracked in-place mutation saves the current model", async () => {
    const document = await parseDocx(await fixtureBuffer());
    expect(getSourceReplayToken(document)).toBeUndefined();
    secondParagraph(document).content.push({
      type: "run",
      content: [{ type: "text", text: " changed" }],
    });
    const buffer = await repackDocx(document, { updateModifiedDate: false });
    const reopened = await parseDocx(buffer);
    expect(paragraphLogicalText(secondParagraph(reopened))).toBe("Second changed");
    expect(await savedXml(buffer)).not.toBe(DOCUMENT_XML);
  });

  test("docx-core operations carry replay identity and serialize only the edited block", async () => {
    const original = await parseDocx(await fixtureBuffer(), { sourceReplay: "tracked" });
    const document = normalizeForOps(original);
    const applied = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
      text: "!",
      runProps: INHERIT_RUN_PROPS,
    });
    if (applied.isErr()) throw applied.error;
    const edited = applied.value.document;
    expect(getSourceReplayToken(edited)).toBe(getSourceReplayToken(original));
    expect(edited.package.document.content.at(0)).toBe(original.package.document.content.at(0));
    expect(secondParagraph(edited)).not.toBe(secondParagraph(original));
    const xml = await savedXml(await repackTracked(edited));
    expect(xml).toContain(UNTOUCHED_BLOCK);
    expect(xml).toContain("!Second");
    expect(paragraphLogicalText(secondParagraph(original))).toBe("Second");
  });
});
