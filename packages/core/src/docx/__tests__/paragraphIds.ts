import { ZipArchive } from "@stll/docx-core/zip";

import { ensureParaIds } from "../ensureParaIds";

/**
 * `xml` without its paragraph ids: a save gives every paragraph one, so a
 * test about anything else a part holds reads the part without them.
 */
export const withoutParagraphIds = (xml: string): string =>
  xml.replace(/\s(?:w14:paraId|w14:textId)="[^"]*"/gu, "");

/**
 * `docx` with an id on every paragraph, as the editor's ingest gives them: a
 * save then adds none, and every part it leaves alone stays byte for byte.
 */
export const withParagraphIds = async (
  docx: Uint8Array | ArrayBuffer,
): Promise<ArrayBuffer> => {
  const { docx: stamped } = await ensureParaIds(docx);
  return stamped.buffer.slice(
    stamped.byteOffset,
    stamped.byteOffset + stamped.byteLength,
  ) as ArrayBuffer;
};

/**
 * `docx` with no paragraph ids in any part, as a producer that writes none
 * leaves it: a save gives every paragraph one, so a test of paragraphs
 * without ids takes them off what it saved.
 */
export const withoutPackageParagraphIds = async (
  docx: Uint8Array | ArrayBuffer,
): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(docx);
  for (const path of zip.paths()) {
    if (!/^word\/[^/]+\.xml$/u.test(path)) continue;
    // oxlint-disable-next-line no-await-in-loop -- each part is read and rewritten in turn
    const xml = await zip.entry(path)?.text();
    if (xml !== undefined) zip.set(path, withoutParagraphIds(xml));
  }
  return (await zip.generate({ compression: "store" })).buffer as ArrayBuffer;
};
