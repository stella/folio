/**
 * Packages that are not WordprocessingML documents, one per reason the CLI
 * refuses a file for, and packages that must still pass. Shared by the CLI
 * and MCP matrices so both surfaces are held to the same list.
 */

import JSZip from "jszip";

import type { InvalidPackageReason } from "../main-document-part";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const OFFICE_DOCUMENT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STRICT_OFFICE_DOCUMENT =
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

type Rewrite = (zip: JSZip) => Promise<void> | void;

const rewritePart = async (zip: JSZip, name: string, edit: (xml: string) => string) => {
  const xml = await zip.file(name)?.async("string");
  if (xml === undefined) throw new Error(`fixture has no ${name}`);
  zip.file(name, edit(xml));
};

const rewritten = async (valid: Uint8Array, rewrite: Rewrite): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(valid);
  await rewrite(zip);
  return await zip.generateAsync({ type: "uint8array" });
};

/** Add a second central-directory record for the same local part. JSZip's name map hides it. */
const duplicateExactMember = (valid: Uint8Array): Uint8Array => {
  const view = new DataView(valid.buffer, valid.byteOffset, valid.byteLength);
  const end = valid.length - 22;
  const count = view.getUint16(end + 10, true);
  const size = view.getUint32(end + 12, true);
  let offset = view.getUint32(end + 16, true);
  let member: { offset: number; length: number } | null = null;
  for (let index = 0; index < count; index += 1) {
    const nameLength = view.getUint16(offset + 28, true);
    const length =
      46 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    const name = new TextDecoder().decode(valid.subarray(offset + 46, offset + 46 + nameLength));
    if (name === "word/document.xml") {
      member = { offset, length };
      break;
    }
    offset += length;
  }
  if (!member) throw new Error("fixture has no main part in its central directory");
  const duplicate = new Uint8Array(valid.length + member.length);
  duplicate.set(valid.subarray(0, end));
  duplicate.set(valid.subarray(member.offset, member.offset + member.length), end);
  duplicate.set(valid.subarray(end), end + member.length);
  const updated = new DataView(duplicate.buffer);
  updated.setUint16(end + member.length + 8, count + 1, true);
  updated.setUint16(end + member.length + 10, count + 1, true);
  updated.setUint32(end + member.length + 12, size + member.length, true);
  return duplicate;
};

export type MalformedCase = {
  name: string;
  reason: InvalidPackageReason;
  build: (valid: Uint8Array) => Promise<Uint8Array> | Uint8Array;
};

export const MALFORMED_PACKAGES: readonly MalformedCase[] = [
  { name: "an empty file", reason: "notZip", build: () => new Uint8Array() },
  {
    name: "junk bytes",
    reason: "notZip",
    build: () => new TextEncoder().encode("not a package ".repeat(64)),
  },
  {
    name: "a truncated archive",
    reason: "notZip",
    build: (valid) => valid.slice(0, Math.floor(valid.length / 2)),
  },
  {
    name: "a ZIP of unrelated entries",
    reason: "contentTypesMissing",
    build: async () => {
      const zip = new JSZip();
      zip.file("unrelated.txt", "hello");
      return await zip.generateAsync({ type: "uint8array" });
    },
  },
  {
    name: "a package without [Content_Types].xml",
    reason: "contentTypesMissing",
    build: (valid) => rewritten(valid, (zip) => void zip.remove("[Content_Types].xml")),
  },
  {
    name: "content types in another namespace",
    reason: "contentTypesInvalid",
    build: (valid) =>
      rewritten(valid, (zip) =>
        rewritePart(zip, "[Content_Types].xml", (xml) =>
          xml.replace("http://schemas.openxmlformats.org/package/2006/content-types", "urn:x"),
        ),
      ),
  },
  {
    name: "a package without _rels/.rels",
    reason: "relationshipsMissing",
    build: (valid) => rewritten(valid, (zip) => void zip.remove("_rels/.rels")),
  },
  {
    name: "package relationships that are not XML",
    reason: "relationshipsInvalid",
    build: (valid) => rewritten(valid, (zip) => void zip.file("_rels/.rels", "<Relationships")),
  },
  {
    name: "no main document relationship",
    reason: "mainRelationshipMissing",
    build: (valid) =>
      rewritten(valid, (zip) =>
        rewritePart(zip, "_rels/.rels", (xml) =>
          xml.replace(OFFICE_DOCUMENT, `${OFFICE_DOCUMENT}-not`),
        ),
      ),
  },
  {
    name: "two main document relationships",
    reason: "mainRelationshipAmbiguous",
    build: (valid) =>
      rewritten(valid, (zip) =>
        rewritePart(zip, "_rels/.rels", (xml) =>
          xml.replace(
            "</Relationships>",
            `<Relationship Id="rId9" Type="${OFFICE_DOCUMENT}" Target="word/document.xml"/></Relationships>`,
          ),
        ),
      ),
  },
  {
    name: "a package without word/document.xml",
    reason: "mainPartMissing",
    build: (valid) => rewritten(valid, (zip) => void zip.remove("word/document.xml")),
  },
  {
    name: "a main part of another content type",
    reason: "mainPartContentType",
    build: (valid) =>
      rewritten(valid, (zip) =>
        rewritePart(zip, "[Content_Types].xml", (xml) =>
          xml.replace(
            MAIN_CONTENT_TYPE,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
          ),
        ),
      ),
  },
  {
    name: "a main part whose root is not a document",
    reason: "mainPartRoot",
    build: (valid) =>
      rewritten(
        valid,
        (zip) => void zip.file("word/document.xml", `<w:styles xmlns:w="${W_NS}"/>`),
      ),
  },
  {
    name: "a document root in another namespace",
    reason: "mainPartRoot",
    build: (valid) =>
      rewritten(
        valid,
        (zip) =>
          void zip.file("word/document.xml", '<w:document xmlns:w="urn:x"><w:body/></w:document>'),
      ),
  },
  {
    name: "a main document folio does not read",
    reason: "mainPartLocation",
    build: (valid) =>
      rewritten(valid, async (zip) => {
        const xml = await zip.file("word/document.xml")?.async("string");
        if (xml === undefined) throw new Error("fixture has no main part");
        zip.remove("word/document.xml");
        zip.file("word/main.xml", xml);
        await rewritePart(zip, "_rels/.rels", (rels) =>
          rels.replace("word/document.xml", "word/main.xml"),
        );
        await rewritePart(zip, "[Content_Types].xml", (types) =>
          types.replace("/word/document.xml", "/word/main.xml"),
        );
      }),
  },
  {
    name: "two entries with one part name",
    reason: "duplicatePartName",
    build: (valid) =>
      rewritten(
        valid,
        (zip) => void zip.file("WORD/DOCUMENT.XML", `<w:document xmlns:w="${W_NS}"/>`),
      ),
  },
  {
    name: "two entries with exactly the same part name",
    reason: "duplicatePartName",
    build: duplicateExactMember,
  },
];

/** Packages that must pass: the check reads namespaces, not prefixes or one profile. */
export const WELL_FORMED_PACKAGES: readonly {
  name: string;
  build: (valid: Uint8Array) => Promise<Uint8Array>;
}[] = [
  { name: "a Transitional package", build: (valid) => Promise.resolve(valid) },
  {
    name: "a Strict package",
    build: (valid) =>
      rewritten(valid, async (zip) => {
        await rewritePart(zip, "_rels/.rels", (xml) =>
          xml.replace(OFFICE_DOCUMENT, STRICT_OFFICE_DOCUMENT),
        );
        await rewritePart(zip, "word/document.xml", (xml) => xml.replaceAll(W_NS, STRICT_W_NS));
        await rewritePart(zip, "word/styles.xml", (xml) => xml.replaceAll(W_NS, STRICT_W_NS));
      }),
  },
  {
    name: "a main part under another prefix, typed by a Default",
    build: (valid) =>
      rewritten(valid, async (zip) => {
        await rewritePart(zip, "word/document.xml", (xml) =>
          xml.replaceAll("<w:", "<x:").replaceAll("</w:", "</x:").replace("xmlns:w=", "xmlns:x="),
        );
        await rewritePart(zip, "[Content_Types].xml", (xml) =>
          xml
            .replace(/<Override PartName="\/word\/document.xml"[^>]*\/>/u, "")
            .replace(
              'Extension="xml" ContentType="application/xml"',
              `Extension="xml" ContentType="${MAIN_CONTENT_TYPE}"`,
            ),
        );
      }),
  },
];

/**
 * Arguments that get each tool past its own validation, so the file is what
 * a command is refused for. Keyed by tool and checked against the registry.
 */
export const TOOL_ARGUMENTS: Readonly<Record<string, Record<string, unknown>>> = {
  read_document: {},
  get_document_outline: {},
  read_section: { handle: "section-1" },
  list_stories: {},
  read_story: { handle: "story-1" },
  find_text: { query: "payment" },
  read_comments: {},
  read_changes: {},
  suggest_changes: {
    operations: [{ type: "insertAfterBlock", blockId: "10000001", text: "Added." }],
  },
  add_comment: { blockId: "10000001", text: "Why?" },
  reply_comment: { commentId: "0", text: "Because." },
  resolve_comment: { commentId: "0" },
  resolve_changes: { all: true },
};
