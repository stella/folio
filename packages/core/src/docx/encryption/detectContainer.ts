/**
 * Detect whether bytes are a normal OOXML ZIP or an OLE/CFB encrypted container.
 *
 * @see MS-OFFCRYPTO — encrypted .docx files are CFB compound files, not ZIP.
 */

export const DOCX_CONTAINER_TYPES = {
  ZIP: "zip",
  CFB: "cfb",
  UNKNOWN: "unknown",
} as const;

export type DocxContainerType = (typeof DOCX_CONTAINER_TYPES)[keyof typeof DOCX_CONTAINER_TYPES];

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04] as const;
const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

const matchesMagic = (bytes: Uint8Array, magic: readonly number[]): boolean => {
  if (bytes.length < magic.length) {
    return false;
  }
  for (let i = 0; i < magic.length; i++) {
    if (bytes[i] !== magic[i]) {
      return false;
    }
  }
  return true;
};

export const detectDocxContainerType = (data: ArrayBuffer | Uint8Array): DocxContainerType => {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (matchesMagic(bytes, ZIP_MAGIC)) {
    return DOCX_CONTAINER_TYPES.ZIP;
  }
  if (matchesMagic(bytes, CFB_MAGIC)) {
    return DOCX_CONTAINER_TYPES.CFB;
  }
  return DOCX_CONTAINER_TYPES.UNKNOWN;
};
