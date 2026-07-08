/**
 * Shared MS-CFB constants and byte helpers for encrypted OOXML containers.
 *
 * @see https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
 */

export const CFB_SIGNATURE = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

export const CFB_HEADER_SIZE = 512;
export const CFB_DIRECTORY_ENTRY_SIZE = 128;
export const CFB_HEADER_DIFAT_ENTRY_COUNT = 109;

export const CFB_FREE_SECTOR = 0xff_ff_ff_ff;
export const CFB_END_OF_CHAIN = 0xff_ff_ff_fe;
export const CFB_DIFAT_SECTOR = 0xff_ff_ff_fc;
export const CFB_NO_STREAM = 0xff_ff_ff_ff;

export const CFB_OBJECT_TYPE = {
  STORAGE: 1,
  STREAM: 2,
  ROOT: 5,
} as const;

export const CFB_BYTE_ORDER = 0xff_fe;
export const CFB_VERSION_3 = 3;
export const CFB_VERSION_4 = 4;
export const CFB_VERSION_3_SECTOR_SIZE = 512;
export const CFB_VERSION_4_SECTOR_SIZE = 4096;
export const CFB_MINI_SECTOR_SHIFT = 6;

export const concatUint8Arrays = (chunks: Uint8Array[]): Uint8Array => {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
};

export const getSectorOffset = (sectorId: number, sectorSize: number): number =>
  (sectorId + 1) * sectorSize;

export const readUint16LE = (view: DataView, offset: number): number =>
  view.getUint16(offset, true);

export const readUint32LE = (view: DataView, offset: number): number =>
  view.getUint32(offset, true);

export const readUint64LE = (view: DataView, offset: number): bigint => {
  const low = BigInt(view.getUint32(offset, true));
  const high = BigInt(view.getUint32(offset + 4, true));
  return (high << 32n) | low;
};

const utf16Decoder = new TextDecoder("utf-16le");

export const decodeUtf16Le = (bytes: Uint8Array): string => {
  const decoded = utf16Decoder.decode(bytes);
  let end = decoded.length;
  while (end > 0 && decoded.charCodeAt(end - 1) === 0) {
    end--;
  }
  return decoded.slice(0, end);
};
