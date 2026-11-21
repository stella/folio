/**
 * The parts a canonical opener rewrote before reading a package, kept so a
 * save writes the package's own bytes for each one the session left alone.
 *
 * A canonical session opens a normalized copy of the package: every story
 * paragraph, content control and revision holds an id of its own. Those ids
 * are minted deterministically, so the same package gets them again on every
 * open, and a save need not write them for a reopen to find them. A save so
 * writes them only into parts it rewrites anyway: a part whose saved text is
 * the text the session read is written as the package held it, and an
 * unedited save changes no part.
 */
import { ZipArchive } from "@stll/docx-core/zip";

import type { Document } from "../types/document";
import type { ReadonlyDocument } from "../types/documentModel";
import { readPackageForSave } from "./mainDocumentPart";

/** One part the opener rewrote: its bytes as the package held them. */
type NormalizedSourcePart = {
  source: Uint8Array<ArrayBuffer>;
  /** SHA-256 of the part as the session read it. */
  readDigest: string;
};

type NormalizedSourceParts = ReadonlyMap<string, NormalizedSourcePart>;

/** By the package buffer the session read, as `Document.originalBuffer`. */
const normalizedSourceParts = new WeakMap<ArrayBuffer, NormalizedSourceParts>();

const sha256Hex = async (bytes: Uint8Array<ArrayBuffer>): Promise<string> => {
  const digest = new Uint8Array(await crypto.subtle.digest("SHA-256", bytes));
  let hex = "";
  for (const byte of digest) hex += byte.toString(16).padStart(2, "0");
  return hex;
};

const isXmlPart = (path: string): boolean => {
  const lower = path.toLowerCase();
  return lower.endsWith(".xml") || lower.endsWith(".rels");
};

const sameBytes = (left: Uint8Array, right: Uint8Array): boolean => {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1)
    if (left[index] !== right[index]) return false;
  return true;
};

type KeepNormalizedSourcePartsOptions = {
  /** The package as the host supplied it. */
  source: ArrayBuffer;
  /** The document read from its normalized copy. */
  document: Document;
};

/**
 * Remember, for `document`'s saves, every XML part its package reads
 * differently from `source`.
 */
export const keepNormalizedSourceParts = async ({
  source,
  document,
}: KeepNormalizedSourcePartsOptions): Promise<void> => {
  const read = document.originalBuffer;
  if (read === undefined || read === source) return;
  // Compare at the writer's part paths so a restored relationship or
  // content type still names the relocated main document.
  const [sourceZip, readZip] = await Promise.all([
    readPackageForSave(source),
    ZipArchive.read(read),
  ]);
  const parts = new Map<string, NormalizedSourcePart>();
  for (const entry of readZip.entries()) {
    if (!isXmlPart(entry.path)) continue;
    const original = sourceZip.entry(entry.path);
    if (original === undefined) continue;
    // oxlint-disable-next-line no-await-in-loop -- one part at a time bounds the bytes held at once
    const [readBytes, sourceBytes] = await Promise.all([
      entry.bytes(),
      original.bytes(),
    ]);
    if (sameBytes(readBytes, sourceBytes)) continue;
    // oxlint-disable-next-line no-await-in-loop -- see above
    const readDigest = await sha256Hex(readBytes);
    parts.set(entry.path, { source: sourceBytes, readDigest });
  }
  if (parts.size > 0) normalizedSourceParts.set(read, parts);
};

/** Carry the kept parts to a copy of the document. */
export const copyNormalizedSourceParts = (
  target: Document,
  source: ReadonlyDocument,
): void => {
  if (!target.originalBuffer || !source.originalBuffer) return;
  const parts = normalizedSourceParts.get(source.originalBuffer);
  if (parts) normalizedSourceParts.set(target.originalBuffer, parts);
};

/**
 * `saved` with the package's own bytes for every kept part `document`'s
 * session left as it read it.
 */
export const withNormalizedSourceParts = async (
  saved: ArrayBuffer,
  document: ReadonlyDocument,
): Promise<ArrayBuffer> => {
  const parts =
    document.originalBuffer === undefined
      ? undefined
      : normalizedSourceParts.get(document.originalBuffer);
  if (parts === undefined) return saved;
  const zip = await ZipArchive.read(saved);
  let restored = 0;
  for (const [path, { source, readDigest }] of parts) {
    const entry = zip.entry(path);
    if (entry === undefined) continue;
    // oxlint-disable-next-line no-await-in-loop -- one part at a time bounds the bytes held at once
    const savedDigest = await sha256Hex(await entry.bytes());
    if (savedDigest !== readDigest) continue;
    zip.set(path, source, { modified: entry.modified });
    restored += 1;
  }
  if (restored === 0) return saved;
  const bytes = await zip.generate({ level: 6 });
  return bytes.buffer;
};
