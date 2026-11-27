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
