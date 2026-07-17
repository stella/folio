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
import { DOCX_MAX_ENTRIES } from "./server/boundedArchive";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

const DIGITAL_SIGNATURE_PARTS = {
  "_xmlsignatures/origin.sigs": "",
  "_xmlsignatures/sig1.xml": '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>',
};

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
  [...xml.matchAll(/w14:paraId=(?<quote>["'])(?<id>[\s\S]*?)\k<quote>/gu)].map(
    (match) => match.groups!["id"]!,
  );

const collectTextIds = (xml: string): string[] =>
  [...xml.matchAll(/w14:textId=(?<quote>["'])(?<id>[\s\S]*?)\k<quote>/gu)].map(
    (match) => match.groups!["id"]!,
  );

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

  test("supports single-quoted ids, namespaces, and mc:Ignorable values", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `${PARA("Kept", " w14:paraId='1A2B3C4D' w14:textId='1A2B3C4D'")}${PARA("Needs one")}`,
        " xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml' xmlns:w15='urn:x' mc:Ignorable='w15'",
      ),
    });

    const result = await ensureParaIds(input);
    const xml = await getPart(result.docx, "word/document.xml");

    expect(result.assigned).toBe(1);
    expect(collectIds(xml)[0]).toBe("1A2B3C4D");
    expect(xml).toContain("mc:Ignorable='w15 w14'");
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
      "word/document.xml": documentXml(
        PARA("Zeroed", ' w14:paraId="00000000" w14:textId="DEADBEEF"'),
      ),
    });

    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(1);

    const xml = await getPart(result.docx, "word/document.xml");
    const id = collectIds(xml)[0];
    expect(id).not.toBe("00000000");
    expect(collectTextIds(xml)[0]).toBe(id);
  });

  test("reassigns later duplicates; the first occurrence keeps the id", async () => {
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `${PARA("Original", ' w14:paraId="ABCD1234"')}${PARA("Copy", ' w14:paraId="ABCD1234" w14:textId="DEADBEEF"')}`,
      ),
    });

    const result = await ensureParaIds(input);
    expect(result.deduplicated).toBe(1);

    const xml = await getPart(result.docx, "word/document.xml");
    const ids = collectIds(xml);
    expect(ids[0]).toBe("ABCD1234");
    expect(ids[1]).not.toBe("ABCD1234");
    expect(ids[1]).toMatch(/^[0-9A-F]{8}$/u);
    expect(collectTextIds(xml)[0]).toBe(ids[1]);
  });

  test("ignores comments, CDATA, and processing instructions during the XML scan", async () => {
    const opaque =
      '<!-- <w:p w14:paraId="AAAA0001"/> --><![CDATA[<w:p w14:paraId="BBBB0002"/>]]><?folio <w:p w14:paraId="CCCC0003"/> ?>';
    const input = await buildDocx({
      "word/document.xml": documentXml(`${opaque}${PARA("Real paragraph")}`),
    });

    const result = await ensureParaIds(input);
    const xml = await getPart(result.docx, "word/document.xml");

    expect(result.assigned).toBe(1);
    expect(collectIds(xml)).toHaveLength(4);
    expect(xml).toContain(opaque);
  });

  test("never modifies paragraphs inside mc:Fallback", async () => {
    const fallbackParagraph = PARA("Shadow copy", ' w14:paraId="FEED0001"');
    const input = await buildDocx({
      "word/document.xml": documentXml(
        `<w:p w14:paraId="FEED0001"><w:r><mc:AlternateContent><mc:Choice Requires="wps">${PARA("Box text")}</mc:Choice><mc:Fallback>${fallbackParagraph}${PARA("No id either")}</mc:Fallback></mc:AlternateContent></w:r></w:p>`,
      ),
    });

    const result = await ensureParaIds(input);
    const xml = await getPart(result.docx, "word/document.xml");

    // The fallback branch is byte-identical: the id-less paragraph stays
    // id-less and the existing id is not treated as a duplicate to rewrite.
    expect(xml).toContain(`<mc:Fallback>${fallbackParagraph}${PARA("No id either")}</mc:Fallback>`);
    // The fallback ID owns its value, so the conflicting outer paragraph is
    // reminted even though it appears first; the mc:Choice paragraph gets an ID.
    expect(result.assigned).toBe(1);
    expect(result.deduplicated).toBe(1);
  });

  test("covers headers, footers, footnotes, and endnotes with document-wide uniqueness", async () => {
    const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W_NS}>${PARA("Header text")}</w:hdr>`;
    const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${W_NS}>${PARA("Footer text")}</w:ftr>`;
    const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W_NS}><w:footnote w:id="1">${PARA("Footnote text")}</w:footnote></w:footnotes>`;
    const endnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes ${W_NS}><w:endnote w:id="1">${PARA("Endnote text")}</w:endnote></w:endnotes>`;
    const input = await buildDocx({
      "word/document.xml": documentXml(PARA("Body")),
      "word/header1.xml": headerXml,
      "word/footer1.xml": footerXml,
      "word/footnotes.xml": footnotesXml,
      "word/endnotes.xml": endnotesXml,
    });

    const result = await ensureParaIds(input);
    expect(result.assigned).toBe(5);

    const bodyIds = collectIds(await getPart(result.docx, "word/document.xml"));
    const headerIds = collectIds(await getPart(result.docx, "word/header1.xml"));
    const footerIds = collectIds(await getPart(result.docx, "word/footer1.xml"));
    const footnoteIds = collectIds(await getPart(result.docx, "word/footnotes.xml"));
    const endnoteIds = collectIds(await getPart(result.docx, "word/endnotes.xml"));
    expect(headerIds).toHaveLength(1);
    expect(footerIds).toHaveLength(1);
    expect(footnoteIds).toHaveLength(1);
    expect(endnoteIds).toHaveLength(1);
    const all = [...bodyIds, ...headerIds, ...footerIds, ...footnoteIds, ...endnoteIds];
    expect(new Set(all).size).toBe(all.length);

    const headerOut = await getPart(result.docx, "word/header1.xml");
    expect(headerOut).toContain('mc:Ignorable="w14"');
  });

  test("leaves the comments part untouched while honoring its ids as taken", async () => {
    const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${W_NS} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0">${PARA("A comment", ' w14:paraId="C0FFEE01"')}</w:comment></w:comments>`;
    const input = await buildDocx({
      "word/document.xml": documentXml(PARA("Body", ' w14:paraId="C0FFEE01"')),
      "word/comments.xml": commentsXml,
    });

    const result = await ensureParaIds(input);
    expect(result.deduplicated).toBe(1);
    const bodyIds = collectIds(await getPart(result.docx, "word/document.xml"));
    expect(bodyIds[0]).not.toBe("C0FFEE01");
    expect(await getPart(result.docx, "word/comments.xml")).toBe(commentsXml);
  });

  test("rejects conflicting w14 or mc namespace bindings before patching", async () => {
    const conflictingRoots = [' xmlns:w14="urn:wrong"', ' xmlns:mc="urn:wrong"'];

    for (const rootAttrs of conflictingRoots) {
      const input = await buildDocx({
        "word/document.xml": documentXml(PARA("Body"), rootAttrs),
      });
      await expect(ensureParaIds(input)).rejects.toThrow(EnsureParaIdsError);
    }
  });

  test("requires explicit opt-in before invalidating package signatures", async () => {
    const input = await buildDocx({
      ...DIGITAL_SIGNATURE_PARTS,
      "word/document.xml": documentXml(PARA("Needs an ID")),
    });

    await expect(ensureParaIds(input)).rejects.toThrow("digitally signed package");

    const result = await ensureParaIds(input, { allowSignedPackageMutation: true });
    expect(result.assigned).toBe(1);
    expect(result.alreadyComplete).toBe(false);
  });

  test("returns an already-complete signed package byte-identically", async () => {
    const input = await buildDocx({
      ...DIGITAL_SIGNATURE_PARTS,
      "word/document.xml": documentXml(
        PARA("Already identified", ' w14:paraId="12345678" w14:textId="12345678"'),
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"',
      ),
    });

    const result = await ensureParaIds(input);

    expect(result.alreadyComplete).toBe(true);
    expect(result.docx).toBe(input);
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

    await expect(ensureParaIds(notDocx)).rejects.toThrow(EnsureParaIdsError);
  });

  test("rejects an untrusted archive over the bounded-load entry-count cap before reading any XML", async () => {
    // Regression guard for routing ensureParaIds through the bounded archive
    // loader (loadDocxArchive): an attacker-supplied ingest buffer with an
    // excessive entry count must fail fast instead of materializing every
    // part's XML with no cap.
    const zip = new JSZip();
    zip.file("[Content_Types].xml", CONTENT_TYPES);
    zip.file("_rels/.rels", ROOT_RELS);
    zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
    zip.file("word/document.xml", documentXml(PARA("Body")));
    for (let i = 0; i < DOCX_MAX_ENTRIES + 5; i++) {
      zip.file(`word/media/filler${i}.bin`, "x");
    }
    const input = await zip.generateAsync({ type: "uint8array" });

    await expect(ensureParaIds(input)).rejects.toThrow(EnsureParaIdsError);
  });

  test("wraps malformed ZIP input in EnsureParaIdsError", async () => {
    const malformed = new Uint8Array([0, 1, 2, 3]);

    const error = await ensureParaIds(malformed).then(
      () => null,
      (reason: unknown) => reason,
    );
    expect(error).toBeInstanceOf(EnsureParaIdsError);
    if (!(error instanceof EnsureParaIdsError)) {
      throw new Error("Expected EnsureParaIdsError");
    }
    expect(error.message).toContain("Failed to normalize paragraph IDs");
    expect(error.cause).toBeInstanceOf(Error);
  });
});
