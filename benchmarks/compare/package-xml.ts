/**
 * Raw OOXML package plumbing for the compare benchmark.
 *
 * The harness builds and edits packages as XML text rather than through
 * folio's document model on purpose. The model is the thing under measurement:
 * generating corpus documents with it would let a parser bug produce inputs the
 * comparison then trivially agrees with, and serializing through it would put
 * engine work inside the setup the timings are supposed to exclude.
 */

import JSZip from "jszip";

/** A package part: XML text, or bytes for media. */
export type PackagePart = string | Uint8Array;

export type DocxPackage = ReadonlyMap<string, PackagePart>;

/**
 * ZIP entry date every generated package is stamped with, so two runs of the
 * generator produce byte-identical corpus bytes.
 */
const FIXED_ZIP_DATE = new Date(Date.UTC(2000, 0, 1));

export const zipPackage = async (parts: DocxPackage): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  for (const name of [...parts.keys()].toSorted()) {
    const part = parts.get(name);
    if (part === undefined) {
      continue;
    }
    zip.file(name, part, { date: FIXED_ZIP_DATE });
  }
  return await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
};

export const unzipPackage = async (buffer: ArrayBuffer): Promise<Map<string, PackagePart>> => {
  const zip = await JSZip.loadAsync(buffer);
  const parts = new Map<string, PackagePart>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) {
      continue;
    }
    parts.set(
      name,
      name.endsWith(".xml") || name.endsWith(".rels")
        ? await file.async("string")
        : await file.async("uint8array"),
    );
  }
  return parts;
};

export const escapeXml = (text: string): string =>
  text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

const BODY_OPEN = "<w:body>";
const BODY_CLOSE = "</w:body>";

type BodyBounds = { start: number; end: number };

const bodyBounds = (documentXml: string): BodyBounds => {
  const start = documentXml.indexOf(BODY_OPEN);
  const end = documentXml.lastIndexOf(BODY_CLOSE);
  if (start === -1 || end === -1) {
    throw new Error("The document part has no w:body.");
  }
  return { start: start + BODY_OPEN.length, end };
};

/**
 * The top-level children of `<w:body>`, as raw XML strings in document order.
 *
 * A regular expression cannot do this: `<w:tbl>` nests `<w:p>`, and a `<w:p>`
 * inside a text box nests further still. The scan tracks element depth, which
 * is enough because no generated element name also appears as text.
 */
export const splitBodyChildren = (documentXml: string): string[] => {
  const { start: bodyStart, end: bodyEnd } = bodyBounds(documentXml);
  const inner = documentXml.slice(bodyStart, bodyEnd);
  const children: string[] = [];
  let depth = 0;
  let start = -1;
  for (const match of inner.matchAll(/<(\/?)[A-Za-z0-9:]+[^>]*?(\/?)>/gu)) {
    const [tag, closing, selfClosing] = match;
    if (closing === "/") {
      depth -= 1;
      if (depth === 0 && start !== -1) {
        children.push(inner.slice(start, match.index + tag.length));
        start = -1;
      }
      continue;
    }
    if (selfClosing === "/") {
      if (depth === 0) {
        children.push(tag);
      }
      continue;
    }
    if (depth === 0) {
      start = match.index;
    }
    depth += 1;
  }
  return children;
};

export const replaceBodyChildren = (documentXml: string, children: readonly string[]): string => {
  const { start, end } = bodyBounds(documentXml);
  return `${documentXml.slice(0, start)}${children.join("")}${documentXml.slice(end)}`;
};

/** Every `<w:t>` value of one block, concatenated. */
export const blockText = (blockXml: string): string =>
  [...blockXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/gu)]
    .map(([, text]) => text ?? "")
    .join("");

/**
 * Rewrite one block's text: the first `<w:t>` takes `text` and the rest are
 * emptied, so a variant changes what a paragraph says without disturbing its
 * run properties, fields, or note references.
 */
export const withBlockText = (blockXml: string, text: string): string => {
  let replaced = false;
  return blockXml.replaceAll(/<w:t(?:\s[^>]*?)?>[\s\S]*?<\/w:t>/gu, (): string => {
    if (replaced) {
      return "<w:t></w:t>";
    }
    replaced = true;
    return `<w:t xml:space="preserve">${escapeXml(text)}</w:t>`;
  });
};

/** The document part of a package, or a throw when the package has none. */
export const documentPartOf = (parts: DocxPackage): string => {
  const part = parts.get("word/document.xml");
  if (typeof part !== "string") {
    throw new Error("The package has no word/document.xml.");
  }
  return part;
};
