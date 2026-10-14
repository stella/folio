import { decodeText } from "../utils/textEncoding";
/** Decode the UTF-8 and UTF-16 encodings XML processors must recognize. */
export const decodeXmlBytes = (bytes: Uint8Array): string => {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return decodeText(bytes.subarray(2), { encoding: "utf-16le" });
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return decodeText(bytes.subarray(2), { encoding: "utf-16be" });
  }
  if (bytes[0] === 0x3c && bytes[1] === 0x00 && bytes[2] === 0x3f && bytes[3] === 0x00) {
    return decodeText(bytes, { encoding: "utf-16le" });
  }
  if (bytes[0] === 0x00 && bytes[1] === 0x3c && bytes[2] === 0x00 && bytes[3] === 0x3f) {
    return decodeText(bytes, { encoding: "utf-16be" });
  }
  const utf8Offset = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  return decodeText(bytes.subarray(utf8Offset), { encoding: "utf-8" });
};
