import JSZip from "jszip";

const MEBIBYTE = 1024 * 1024;

const CENTRAL_DIRECTORY_SIGNATURE = 0x02_01_4b_50;
const LOCAL_HEADER_SIGNATURE = 0x04_03_4b_50;

/**
 * A small archive whose `entryPath` inflates to `inflatedMebibytes` of one
 * repeated byte. The compressed archive stays a few kilobytes.
 */
export const compressibleArchive = async ({
  entryPath,
  inflatedMebibytes,
}: {
  entryPath: string;
  inflatedMebibytes: number;
}): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("_rels/.rels", "<Relationships/>");
  zip.file("word/document.xml", "<w:document/>");
  zip.file(entryPath, "x".repeat(inflatedMebibytes * MEBIBYTE));
  return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

/** A small archive holding `entryCount` tiny entries. */
export const manyEntryArchive = async (entryCount: number): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("word/document.xml", "<w:document/>");
  for (let index = 0; index < entryCount; index += 1) {
    zip.file(`customXml/item${String(index)}.xml`, "<a/>", { createFolders: false });
  }
  return await zip.generateAsync({ type: "uint8array" });
};

/**
 * Rewrite the uncompressed size `entryPath` declares, in both its local header
 * and its central-directory record, leaving the compressed data untouched. The
 * archive then states a size its content does not have.
 */
export const withDeclaredSize = (
  source: Uint8Array,
  entryPath: string,
  declaredBytes: number,
): Uint8Array => {
  const bytes = source.slice();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const name = new TextEncoder().encode(entryPath);
  const nameAt = (offset: number): boolean =>
    offset + name.length <= bytes.length &&
    name.every((byte, index) => bytes[offset + index] === byte);

  let patched = 0;
  for (let offset = 0; offset + 46 <= bytes.length; offset += 1) {
    const signature = view.getUint32(offset, true);
    if (
      signature === CENTRAL_DIRECTORY_SIGNATURE &&
      view.getUint16(offset + 28, true) === name.length &&
      nameAt(offset + 46)
    ) {
      view.setUint32(offset + 24, declaredBytes, true);
      patched += 1;
    } else if (
      signature === LOCAL_HEADER_SIGNATURE &&
      view.getUint16(offset + 26, true) === name.length &&
      nameAt(offset + 30)
    ) {
      view.setUint32(offset + 22, declaredBytes, true);
      patched += 1;
    }
  }
  if (patched !== 2) {
    throw new Error(`Expected one local and one central record for ${entryPath}`);
  }
  return bytes;
};

/**
 * An archive whose `entryPath` inflates to several mebibytes while declaring
 * 64 bytes: under every size and ratio bound until it is actually inflated.
 */
export const understatedEntryArchive = async (entryPath: string): Promise<Uint8Array> =>
  withDeclaredSize(await compressibleArchive({ entryPath, inflatedMebibytes: 3 }), entryPath, 64);
