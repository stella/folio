import type JSZip from "jszip";
import { encodeUtf8 } from "./encoding";

/** Encode complete text before JSZip divides the entry into byte chunks. */
export const writeZipEntry = (
  zip: JSZip,
  path: string,
  content: string | Uint8Array | ArrayBuffer | Promise<Uint8Array>,
  options?: JSZip.JSZipFileOptions,
): JSZip =>
  zip.file(path, typeof content === "string" ? encodeUtf8(content) : content, {
    ...options,
    binary: true,
    base64: false,
  });
