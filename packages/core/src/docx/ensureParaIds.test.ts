/**
 * Unit tests for the headless `ensureParaIds` ingest pass.
 *
 * Fixtures are built in-test as minimal WordprocessingML packages in the
 * shape python-docx / Google Docs exports produce: no `w14` namespace, no
 * `mc:Ignorable`, no paraIds anywhere.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { isSequentialFolioBlockId } from "../types/block-id";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ensureParaIds, EnsureParaIdsError } from "./ensureParaIds";
import { parseDocx } from "./parser";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const documentXml = (body: string, extraRootAttrs = ""): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}${extraRootAttrs}><w:body>${body}<w:sectPr/></w:body></w:document>`;

const buildDocx = async (parts: Record<string, string>): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  for (const [path, content] of Object.entries(parts)) {
    zip.file(path, content);
  }
  return zip.generateAsync({ type: "uint8array" });
};

const getPart = async (docx: Uint8Array, path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(docx);
  const entry = zip.file(path);
  if (!entry) {
    throw new Error(`missing part: ${path}`);
  }
  return entry.async("text");
};

const collectIds = (xml: string): string[] =>
  [...xml.matchAll(/w14:paraId="(?<id>[^"]*)"/gu)].map((match) => match.groups!["id"]!);

const PARA = (text: string, attrs = ""): string =>
  `<w:p${attrs}><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("ensureParaIds", () => {
  test("backfills paraId + textId on every paragraph, including table cells", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `${PARA("First")}<w:p/><w:tbl><w:tr><w:tc><w:tcPr/>${PARA("Cell")}</w:tc></w:tr></w:tbl>`,
      ),
    });

    const result = await ensureParaIds(input);
    expect(result.alreadyComplete).toBe(false);
    expect(result.assigned).toBe(3);
    expect(result.deduplicated).toBe(0);

    const xml = await getPart(result.docx, "word/document.xml");
    const ids = collectIds(xml);
    expect(ids).toHaveLength(3);
    for (const id of ids) {
      expect(id).toMatch(/^[0-9A-F]{8}$/u);
      expect(id).not.toBe("00000000");
    }
    expect(new Set(ids).size).toBe(3);
    // textId minted alongside, same value.
    for (const id of ids) {
      expect(xml).toContain(`w14:paraId="${id}" w14:textId="${id}"`);
    }
  });

  test("declares xmlns:w14 / xmlns:mc and mc:Ignorable on a root that lacks them", async () => {
    const input = await buildDocx({ "word/document.xml": documentXml(PARA("Hello")) });

    const { docx } = await ensureParaIds(input);
    const xml = await getPart(docx, "word/document.xml");

    expect(xml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(xml).toContain('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"');
    expect(xml).toContain('mc:Ignorable="w14"');
    // Declarations land on the root element, before its first child.
    expect(xml.indexOf("xmlns:w14")).toBeLessThan(xml.indexOf("<w:body>"));
  });

  test("appends w14 to an existing mc:Ignorable and keeps declared namespaces single", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        PARA("Hello"),
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="urn:x" mc:Ignorable="w15"',
      ),
    });

    const { docx } = await ensureParaIds(input);
    const xml = await getPart(docx, "word/document.xml");

    expect(xml).toContain('mc:Ignorable="w15 w14"');
    expect(xml.match(/xmlns:mc=/gu)).toHaveLength(1);
    expect(xml.match(/xmlns:w14=/gu)).toHaveLength(1);
  });

  test("is deterministic: same input bytes produce identical output bytes", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(`${PARA("One")}${PARA("Two")}`),
    });

    const first = await ensureParaIds(input);
    const second = await ensureParaIds(input);
    expect(first.docx).toEqual(second.docx);
  });

  test("is idempotent: a normalized document short-circuits to the input bytes", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(`${PARA("One")}${PARA("Two")}`),
    });

    const first = await ensureParaIds(input);
    const second = await ensureParaIds(first.docx);

    expect(second.alreadyComplete).toBe(true);
    expect(second.assigned).toBe(0);
    expect(second.deduplicated).toBe(0);
    // Byte-identical by construction: the very same buffer comes back.
    expect(second.docx).toBe(first.docx);
  });

  test("preserves existing ids verbatim and only fills gaps", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `${PARA("Kept", ' w14:paraId="1A2B3C4D" w14:textId="1A2B3C4D"')}${PARA("Needs one")}`,
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"',
      ),
    });

    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(1);

    const xml = await getPart(result.docx, "word/document.xml");
    const ids = collectIds(xml);
    expect(ids[0]).toBe("1A2B3C4D");
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(ids[1]).not.toBe("1A2B3C4D");
  });

  test("treats the reserved all-zero paraId as unassigned", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(PARA("Zeroed", ' w14:paraId="00000000"')),
    });

    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(1);

    const xml = await getPart(result.docx, "word/document.xml");
    expect(collectIds(xml)[0]).not.toBe("00000000");
  });

  test("reassigns later duplicates; the first occurrence keeps the id", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `${PARA("Original", ' w14:paraId="ABCD1234"')}${PARA("Copy", ' w14:paraId="ABCD1234"')}`,
      ),
    });

    const result = await ensureParaIds(input);
    expect(result.deduplicated).toBe(1);

    const xml = await getPart(result.docx, "word/document.xml");
    const ids = collectIds(xml);
    expect(ids[0]).toBe("ABCD1234");
    expect(ids[1]).not.toBe("ABCD1234");
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
  });

  test("never modifies paragraphs inside mc:Fallback", async () => {
    const fallbackParagraph = PARA("Shadow copy", ' w14:paraId="FEED0001"');
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">${PARA("Box text")}</mc:Choice><mc:Fallback>${fallbackParagraph}${PARA("No id either")}</mc:Fallback></mc:AlternateContent></w:r></w:p>`,
      ),
    });

    const result = await ensureParaIds(input);
    const xml = await getPart(result.docx, "word/document.xml");

    // The fallback branch is byte-identical: the id-less paragraph stays
    // id-less and the existing id is not treated as a duplicate to rewrite.
    expect(xml).toContain(`<mc:Fallback>${fallbackParagraph}${PARA("No id either")}</mc:Fallback>`);
    // The outer paragraph and the mc:Choice paragraph both got ids.
    expect(result.assigned).toBe(2);
  });

  test("covers headers and footnotes with document-wide uniqueness", async () => {
    const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W_NS}>${PARA("Header text")}</w:hdr>`;
    const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W_NS}><w:footnote w:id="1">${PARA("Footnote text")}</w:footnote></w:footnotes>`;
    const input = await buildDocx({
      "word/document.xml": documentXml(PARA("Body")),
      "word/header1.xml": headerXml,
      "word/footnotes.xml": footnotesXml,
    });

    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(3);

    const bodyIds = collectIds(await getPart(result.docx, "word/document.xml"));
    const headerIds = collectIds(await getPart(result.docx, "word/header1.xml"));
    const footnoteIds = collectIds(await getPart(result.docx, "word/footnotes.xml"));
    expect(headerIds).toHaveLength(1);
    expect(footnoteIds).toHaveLength(1);
    const all = [...bodyIds, ...headerIds, ...footnoteIds];
    expect(new Set(all).size).toBe(all.length);

    const headerOut = await getPart(result.docx, "word/header1.xml");
    expect(headerOut).toContain('mc:Ignorable="w14"');
  });

  test("leaves the comments part untouched while honoring its ids as taken", async () => {
    const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${W_NS}><w:comment w:id="0"><w:p><w:r><w:t>A comment</w:t></w:r></w:p></w:comment></w:comments>`;
    const input = await buildDocx({
      "word/document.xml": documentXml(PARA("Body")),
      "word/comments.xml": commentsXml,
    });

    const result = await ensureParaIds(input);
    expect(await getPart(result.docx, "word/comments.xml")).toBe(commentsXml);
  });

  test("a normalized document snapshots with zero seq- block ids", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `${PARA("First paragraph")}${PARA("Second paragraph")}<w:tbl><w:tr><w:tc><w:tcPr/>${PARA("Cell paragraph")}</w:tc></w:tr></w:tbl>`,
      ),
    });

    const { docx } = await ensureParaIds(input);
    const buffer = docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength);
    const parsed = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
    const snapshot = createFolioAIEditSnapshot(toProseDoc(parsed));

    expect(snapshot.blocks.length).toBeGreaterThanOrEqual(3);
    for (const block of snapshot.blocks) {
      expect(isSequentialFolioBlockId(block.id)).toBe(false);
    }
  });

  test("rejects a buffer without word/document.xml", async () => {
    const zip = new JSZip();
    zip.file("hello.txt", "not a docx");
    const notDocx = await zip.generateAsync({ type: "uint8array" });

    expect(ensureParaIds(notDocx)).rejects.toThrow(EnsureParaIdsError);
  });
});
