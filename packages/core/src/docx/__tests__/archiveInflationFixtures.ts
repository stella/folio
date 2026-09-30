import JSZip from "jszip";

const MEBIBYTE = 1024 * 1024;

const CENTRAL_DIRECTORY_SIGNATURE = 0x02_01_4b_50;
const LOCAL_HEADER_SIGNATURE = 0x04_03_4b_50;
const DEFLATE_METHOD = 8;

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

type EntryHeaderPatch = {
  /** Uncompressed size to declare. */
  declaredBytes?: number;
  /** Compression method to declare. */
  method?: number;
};

/**
 * Rewrite what `entryPath` declares, in both its local header and its
 * central-directory record, leaving its data untouched. The archive then
 * states a size, or a method, its content does not have.
 */
export const withEntryHeaders = (
  source: Uint8Array,
  entryPath: string,
  { declaredBytes, method }: EntryHeaderPatch,
): Uint8Array => {
  const bytes = source.slice();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const name = new TextEncoder().encode(entryPath);
  const nameAt = (offset: number): boolean =>
    offset + name.length <= bytes.length &&
    name.every((byte, index) => bytes[offset + index] === byte);
  const patchHeader = (methodOffset: number, sizeOffset: number) => {
    if (method !== undefined) {
      view.setUint16(methodOffset, method, true);
    }
    if (declaredBytes !== undefined) {
      view.setUint32(sizeOffset, declaredBytes, true);
    }
  };

  let patched = 0;
  for (let offset = 0; offset + 46 <= bytes.length; offset += 1) {
    const signature = view.getUint32(offset, true);
    if (
      signature === CENTRAL_DIRECTORY_SIGNATURE &&
      view.getUint16(offset + 28, true) === name.length &&
      nameAt(offset + 46)
    ) {
      patchHeader(offset + 10, offset + 24);
      patched += 1;
    } else if (
      signature === LOCAL_HEADER_SIGNATURE &&
      view.getUint16(offset + 26, true) === name.length &&
      nameAt(offset + 30)
    ) {
      patchHeader(offset + 8, offset + 22);
      patched += 1;
    }
  }
  if (patched !== 2) {
    throw new Error(`Expected one local and one central record for ${entryPath}`);
  }
  return bytes;
};

/** Rewrite the uncompressed size `entryPath` declares. */
export const withDeclaredSize = (
  source: Uint8Array,
  entryPath: string,
  declaredBytes: number,
): Uint8Array => withEntryHeaders(source, entryPath, { declaredBytes });

/**
 * An archive whose `entryPath` inflates to several mebibytes while declaring
 * 64 bytes: under every size and ratio bound until it is actually inflated.
 */
export const understatedEntryArchive = async (entryPath: string): Promise<Uint8Array> =>
  withDeclaredSize(await compressibleArchive({ entryPath, inflatedMebibytes: 3 }), entryPath, 64);

/** Writes bits least-significant first, the order DEFLATE packs them in. */
class BitWriter {
  private readonly bytes: number[] = [];
  private current = 0;
  private used = 0;

  bits(value: number, count: number): void {
    for (let index = 0; index < count; index += 1) {
      this.current |= ((value >> index) & 1) << this.used;
      this.used += 1;
      if (this.used === 8) {
        this.bytes.push(this.current);
        this.current = 0;
        this.used = 0;
      }
    }
  }

  /** A Huffman code, which DEFLATE packs most-significant bit first. */
  code(value: number, length: number): void {
    for (let index = length - 1; index >= 0; index -= 1) {
      this.bits((value >> index) & 1, 1);
    }
  }

  finish(): Uint8Array {
    if (this.used !== 0) {
      throw new Error("Block does not end on a byte boundary");
    }
    return Uint8Array.from(this.bytes);
  }
}

const CODE_LENGTH_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

/**
 * One non-final dynamic-Huffman block: a zero byte, then `matches` copies of
 * 258 bytes at distance 1. The block's codes give the length-258 symbol and
 * the distance symbol one bit each, the densest output DEFLATE allows, and a
 * multiple of four matches keeps the block byte-aligned so it can repeat.
 */
const zeroRunBlock = (matches: number): Uint8Array => {
  const writer = new BitWriter();
  writer.bits(0, 1); // not the final block
  writer.bits(2, 2); // dynamic Huffman codes
  writer.bits(286 - 257, 5); // literal/length codes
  writer.bits(2 - 1, 5); // distance codes
  writer.bits(19 - 4, 4); // code-length codes
  // Code-length alphabet: 18 (a run of zeros) = "0", 1 = "10", 2 = "11".
  const codeLengthLengths = new Map([
    [18, 1],
    [1, 2],
    [2, 2],
  ]);
  for (const symbol of CODE_LENGTH_ORDER) {
    writer.bits(codeLengthLengths.get(symbol) ?? 0, 3);
  }
  // Literal/length lengths: 0 -> 2, 1..255 -> 0, 256 -> 2, 257..284 -> 0, 285 -> 1.
  writer.code(0b11, 2);
  writer.code(0, 1);
  writer.bits(138 - 11, 7);
  writer.code(0, 1);
  writer.bits(117 - 11, 7);
  writer.code(0b11, 2);
  writer.code(0, 1);
  writer.bits(28 - 11, 7);
  writer.code(0b10, 2);
  // Distance lengths: 0 -> 1, 1 -> 1.
  writer.code(0b10, 2);
  writer.code(0b10, 2);
  // Canonical codes: 285 = "0", literal 0 = "10", end of block = "11"; distance 0 = "0".
  writer.code(0b10, 2);
  for (let index = 0; index < matches; index += 1) {
    writer.code(0, 1);
    writer.code(0, 1);
  }
  writer.code(0b11, 2);
  return writer.finish();
};

/** An empty final block with fixed codes. */
const FINAL_BLOCK = [0x03, 0x00];

/**
 * A raw DEFLATE stream of zero bytes, about 1032 bytes out per byte in.
 * `inflatedMebibytes` is approximate (each block adds one byte).
 */
export const rawDeflateOfZeros = (inflatedMebibytes: number): Uint8Array => {
  const matchesPerBlock = 4096;
  const block = zeroRunBlock(matchesPerBlock);
  const blockOutput = 1 + 258 * matchesPerBlock;
  const blocks = Math.ceil((inflatedMebibytes * MEBIBYTE) / blockOutput);
  const stream = new Uint8Array(block.length * blocks + FINAL_BLOCK.length);
  for (let index = 0; index < blocks; index += 1) {
    stream.set(block, index * block.length);
  }
  stream.set(FINAL_BLOCK, block.length * blocks);
  return stream;
};

/**
 * An archive holding `entryPath` as the given raw DEFLATE data, declaring
 * `declaredBytes`. The data is written stored and then relabelled as
 * compressed, so JSZip never has to produce it.
 */
export const archiveWithRawDeflate = async ({
  entryPath,
  deflated,
  declaredBytes,
}: {
  entryPath: string;
  deflated: Uint8Array;
  declaredBytes: number;
}): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("_rels/.rels", "<Relationships/>");
  zip.file("word/document.xml", "<w:document/>");
  zip.file(entryPath, deflated, { binary: true, compression: "STORE" });
  const stored = await zip.generateAsync({ type: "uint8array" });
  return withEntryHeaders(stored, entryPath, { declaredBytes, method: DEFLATE_METHOD });
};
