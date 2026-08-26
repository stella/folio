import { FolioDocxReviewer } from "../../ai-edits/headless";
import { toArrayBuffer } from "../../utils/docxInput";
import { ensureParaIds } from "../ensureParaIds";
import { parseDocx } from "../parser";
import { createDocx } from "../rezip";
import {
  type BilingualRow,
  type CreateBilingualDocumentOptions,
  createBilingualDocument,
  readBilingualDocument,
} from "./createBilingualDocument";

export type CreateBilingualDocxOptions = Omit<
  CreateBilingualDocumentOptions,
  "editableParagraphIds"
>;

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
  options: CreateBilingualDocxOptions,
): Promise<CreateBilingualDocxResult> {
  const stamped = await ensureParaIds(input);
  const stampedBuffer = await toArrayBuffer(stamped.docx);
  const editableParagraphIds = new Set(
    (await FolioDocxReviewer.fromBuffer(stampedBuffer)).snapshot().blocks.map(({ id }) => id),
  );
  const source = await parseDocx(stampedBuffer, { preloadFonts: false });
  const { document, warnings } = createBilingualDocument(source, {
    ...options,
    editableParagraphIds,
  });
  const buffer = await createDocx(document);
  const rows = await readBilingualDocx(buffer);
  return { buffer, rows, warnings };
}

/** Bytes-in form of {@link readBilingualDocument}. */
export async function readBilingualDocx(input: ArrayBuffer | Uint8Array): Promise<BilingualRow[]> {
  const buffer = await toArrayBuffer(input);
  const document = await parseDocx(buffer, { preloadFonts: false });
  const editableParagraphIds = new Set(
    (await FolioDocxReviewer.fromBuffer(buffer)).snapshot().blocks.map(({ id }) => id),
  );
  return readBilingualDocument(document, editableParagraphIds);
}
