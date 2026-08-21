import { ensureParaIds } from "../ensureParaIds";
import { parseDocx } from "../parser";
import { createDocx } from "../rezip";
import {
  type BilingualRow,
  type CreateBilingualDocumentOptions,
  createBilingualDocument,
  readBilingualDocument,
} from "./createBilingualDocument";

export type CreateBilingualDocxResult = {
  buffer: ArrayBuffer;
  rows: BilingualRow[];
  warnings: string[];
};

/**
 * Bytes-in / bytes-out form of {@link createBilingualDocument}: stamp every
 * paragraph with a `paraId` (so left-column rows are addressable later), parse
 * the DOCX (no font preloading, so it never touches the DOM), lay the body out
 * as a two-column table, and repack onto the original package so theme, fonts,
 * media, headers and footers carry over untouched.
 */
export async function createBilingualDocx(
  input: ArrayBuffer | Uint8Array,
  options: CreateBilingualDocumentOptions,
): Promise<CreateBilingualDocxResult> {
  const stamped = await ensureParaIds(input);
  const source = await parseDocx(stamped.docx, { preloadFonts: false });
  const { document, rows, warnings } = createBilingualDocument(source, options);
  const buffer = await createDocx(document);
  return { buffer, rows, warnings };
}

/** Bytes-in form of {@link readBilingualDocument}. */
export async function readBilingualDocx(input: ArrayBuffer | Uint8Array): Promise<BilingualRow[]> {
  const document = await parseDocx(input, { preloadFonts: false });
  return readBilingualDocument(document);
}
