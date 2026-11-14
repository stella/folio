/**
 * A package whose main document sits at another path than
 * `word/document.xml`, as some producers write it, and the check that a saved
 * package holds exactly one main document.
 */

import { expect } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { createEmptyDocument } from "../../utils/createDocument";
import { mainDocumentPartPath } from "../mainDocumentPart";
import { parseDocx } from "../parser";
import { createDocx } from "../rezip";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
export const DOCUMENT = `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:t>Authored text</w:t></w:r></w:p></w:body></w:document>`;
export const MAIN_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/** The standard package, with its main document moved to `path`. */
export const packageWithMainAt = async (
  path: string,
  target = path,
): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  const relationships = await zip.entry("word/_rels/document.xml.rels")?.text();
  const packageRelationships = await zip.entry("_rels/.rels")?.text();
  const contentTypes = await zip.entry("[Content_Types].xml")?.text();
  if (!relationships || !packageRelationships || !contentTypes) {
    throw new Error("The base package is missing a part");
  }
  zip.delete("word/document.xml");
  zip.delete("word/_rels/document.xml.rels");
  zip.set(path, DOCUMENT);
  const slash = path.lastIndexOf("/");
  zip.set(
    `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`,
    relationships,
  );
  zip.set(
    "_rels/.rels",
    packageRelationships.replace(
      /Target="\/?word\/document\.xml"/u,
      `Target="${target}"`,
    ),
  );
  zip.set(
    "[Content_Types].xml",
    contentTypes.replace(
      /PartName="\/word\/document\.xml"/u,
      `PartName="/${path}"`,
    ),
  );
  return (await zip.generate()).slice().buffer;
};

/** One main part, named by the root relationships and typed by the content types. */
export const expectOneMainDocument = async (
  bytes: ArrayBuffer | null,
  text: string,
): Promise<void> => {
  if (bytes === null) throw new Error("The save was refused");
  const saved = await ZipArchive.read(bytes);
  expect(await mainDocumentPartPath(saved)).toBe("word/document.xml");
  expect(saved.has("word/document2.xml")).toBe(false);
  expect(saved.has("word/_rels/document2.xml.rels")).toBe(false);
  const contentTypes = (await saved.entry("[Content_Types].xml")?.text()) ?? "";
  const mainOverrides = [
    ...contentTypes.matchAll(/<Override\b[^>]*>/gu),
  ].filter(([element]) => element.includes(MAIN_TYPE));
  expect(mainOverrides).toHaveLength(1);
  expect(mainOverrides[0]?.[0]).toContain('PartName="/word/document.xml"');
  const reopened = await FolioDocxReviewer.fromBuffer(bytes);
  expect(reopened.getContent().map((block) => block.text)).toEqual([text]);
};

/** The package as its producer wrote it, and the document read from it holding those bytes. */
export const openedAsWritten = async () => {
  const raw = await packageWithMainAt("word/document2.xml");
  const document = await parseDocx(raw.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return { raw, document: { ...document, originalBuffer: raw } };
};
