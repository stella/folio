import { parseDocx } from "../parser";
import { createDocx } from "../rezip";
import {
  type BilingualRow,
  type CreateBilingualDocumentOptions,
  createBilingualDocument,
} from "./createBilingualDocument";

export type CreateBilingualDocxResult = {
  buffer: ArrayBuffer;
  rows: BilingualRow[];
  warnings: string[];
};

/**
 * Bytes-in / bytes-out form of {@link createBilingualDocument}: parse the DOCX
 * (no font preloading, so it never touches the DOM), lay the body out as a
 * two-column table, and repack onto the original package so theme, fonts,
 * media, headers and footers carry over untouched.
 */
export async function createBilingualDocx(
  input: ArrayBuffer | Uint8Array,
  options: CreateBilingualDocumentOptions,
): Promise<CreateBilingualDocxResult> {
  const source = await parseDocx(input, { preloadFonts: false });
  const { document, rows, warnings } = createBilingualDocument(source, options);
  const buffer = await createDocx(document);
  return { buffer, rows, warnings };
}
