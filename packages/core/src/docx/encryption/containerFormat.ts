/**
 * OOXML container format detection (ZIP vs encrypted OLE compound file).
 *
 * @see https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
 */

export const DOCX_CONTAINER_TYPES = {
  ZIP: "zip",
  CFB: "cfb",
  UNKNOWN: "unknown",
} as const;

export type DocxContainerType = (typeof DOCX_CONTAINER_TYPES)[keyof typeof DOCX_CONTAINER_TYPES];

const ZIP_SIGNATURE = [0x50, 0x4b, 0x03, 0x04] as const;
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

const startsWith = (bytes: Uint8Array, signature: readonly number[]): boolean => {
  if (bytes.length < signature.length) {
    return false;
  }
  for (let i = 0; i < signature.length; i++) {
    if (bytes[i] !== signature[i]) {
      return false;
    }
  }
  return true;
};

export const detectDocxContainerType = (data: ArrayBuffer | Uint8Array): DocxContainerType => {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (startsWith(bytes, ZIP_SIGNATURE)) {
    return DOCX_CONTAINER_TYPES.ZIP;
  }
  if (startsWith(bytes, OLE_SIGNATURE)) {
    return DOCX_CONTAINER_TYPES.CFB;
  }
  return DOCX_CONTAINER_TYPES.UNKNOWN;
};
