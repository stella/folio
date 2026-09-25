/**
 * Minimal sfnt (TrueType/OpenType) face index for local reference fonts.
 * Only the table directory, `name`, `OS/2`, and `head` records are read, so a
 * directory of large font files can be indexed without loading their glyphs.
 * Font files are read at runtime only and never copied.
 */

import { open, readdir } from "node:fs/promises";
import path from "node:path";

import { normalizeFontName } from "./fontNames";

export type FontFaceRecord = {
  filePath: string;
  /** PostScript name (`name` ID 6), as PDFs report embedded faces. */
  postscriptName: string;
  /** Typographic family (`name` ID 16), else the legacy family (ID 1). */
  family: string;
  weight: number;
  style: "normal" | "italic";
};

const SFNT_VERSIONS = new Set([0x0001_0000, 0x4f54_544f /* OTTO */, 0x7472_7565 /* true */]);
const TABLE_DIRECTORY_OFFSET = 12;
const TABLE_RECORD_SIZE = 16;
const NAME_RECORD_SIZE = 12;
const NAME_FAMILY = 1;
const NAME_POSTSCRIPT = 6;
const NAME_TYPOGRAPHIC_FAMILY = 16;
const PLATFORM_UNICODE = 0;
const PLATFORM_MACINTOSH = 1;
const PLATFORM_WINDOWS = 3;
const WINDOWS_ENGLISH_US = 0x0409;
const OS2_WEIGHT_OFFSET = 4;
const OS2_FS_SELECTION_OFFSET = 62;
const FS_SELECTION_ITALIC = 1;
const FS_SELECTION_OBLIQUE = 1 << 9;
const HEAD_MAC_STYLE_OFFSET = 44;
const MAC_STYLE_BOLD = 1;
const MAC_STYLE_ITALIC = 2;
const DEFAULT_WEIGHT = 400;
const BOLD_WEIGHT = 700;
const MAX_TABLES = 512;
const INDEXED_EXTENSIONS = new Set([".ttf", ".otf"]);

type TableSlice = { offset: number; length: number };
type ReadAt = (offset: number, length: number) => Promise<Uint8Array>;

const decodeName = (platformId: number, bytes: Uint8Array): string => {
  if (platformId === PLATFORM_MACINTOSH) {
    return String.fromCharCode(...bytes);
  }
  let text = "";
  for (let index = 0; index + 1 < bytes.length; index += 2) {
    text += String.fromCharCode(((bytes[index] ?? 0) << 8) | (bytes[index + 1] ?? 0));
  }
  return text;
};

/** Preference of a `name` record: Windows US English, then any Windows or
 * Unicode record, then Macintosh Roman. Lower is better. */
const nameRecordRank = (platformId: number, languageId: number): number | undefined => {
  if (platformId === PLATFORM_WINDOWS) return languageId === WINDOWS_ENGLISH_US ? 0 : 1;
  if (platformId === PLATFORM_UNICODE) return 1;
  if (platformId === PLATFORM_MACINTOSH) return 2;
  return undefined;
};

const readNames = async (readAt: ReadAt, table: TableSlice): Promise<Map<number, string>> => {
  const bytes = await readAt(table.offset, table.length);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.byteLength < 6) return new Map();
  const count = view.getUint16(2);
  const storageOffset = view.getUint16(4);
  const best = new Map<number, { rank: number; value: string }>();
  for (let index = 0; index < count; index += 1) {
    const record = 6 + index * NAME_RECORD_SIZE;
    if (record + NAME_RECORD_SIZE > bytes.byteLength) break;
    const platformId = view.getUint16(record);
    const languageId = view.getUint16(record + 4);
    const nameId = view.getUint16(record + 6);
    if (
      nameId !== NAME_FAMILY &&
      nameId !== NAME_POSTSCRIPT &&
      nameId !== NAME_TYPOGRAPHIC_FAMILY
    ) {
      continue;
    }
    const rank = nameRecordRank(platformId, languageId);
    if (rank === undefined || (best.get(nameId)?.rank ?? Number.POSITIVE_INFINITY) <= rank) {
      continue;
    }
    const length = view.getUint16(record + 8);
    const start = storageOffset + view.getUint16(record + 10);
    if (start + length > bytes.byteLength) continue;
    const value = decodeName(platformId, bytes.subarray(start, start + length)).trim();
    if (value.length > 0) best.set(nameId, { rank, value });
  }
  return new Map([...best].map(([nameId, { value }]) => [nameId, value]));
};

/** Parse one face from an sfnt file through a positional reader. Returns
 * undefined for collections, other formats, or faces without usable names. */
export const parseFontFace = async (
  filePath: string,
  readAt: ReadAt,
): Promise<FontFaceRecord | undefined> => {
  const header = await readAt(0, TABLE_DIRECTORY_OFFSET);
  if (header.byteLength < TABLE_DIRECTORY_OFFSET) return undefined;
  const headerView = new DataView(header.buffer, header.byteOffset, header.byteLength);
  if (!SFNT_VERSIONS.has(headerView.getUint32(0))) return undefined;
  const numTables = headerView.getUint16(4);
  if (numTables === 0 || numTables > MAX_TABLES) return undefined;

  const directory = await readAt(TABLE_DIRECTORY_OFFSET, numTables * TABLE_RECORD_SIZE);
  const directoryView = new DataView(directory.buffer, directory.byteOffset, directory.byteLength);
  const tables = new Map<string, TableSlice>();
  for (let index = 0; index < numTables; index += 1) {
    const record = index * TABLE_RECORD_SIZE;
    if (record + TABLE_RECORD_SIZE > directory.byteLength) break;
    const tag = String.fromCharCode(...directory.subarray(record, record + 4));
    tables.set(tag, {
      offset: directoryView.getUint32(record + 8),
      length: directoryView.getUint32(record + 12),
    });
  }

  const nameTable = tables.get("name");
  if (!nameTable) return undefined;
  const names = await readNames(readAt, nameTable);
  const postscriptName = names.get(NAME_POSTSCRIPT);
  const family = names.get(NAME_TYPOGRAPHIC_FAMILY) ?? names.get(NAME_FAMILY);
  if (postscriptName === undefined || family === undefined) return undefined;

  let weight: number | undefined;
  let italic: boolean | undefined;
  const os2 = tables.get("OS/2");
  if (os2 && os2.length >= OS2_FS_SELECTION_OFFSET + 2) {
    const bytes = await readAt(os2.offset, OS2_FS_SELECTION_OFFSET + 2);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const weightClass = view.getUint16(OS2_WEIGHT_OFFSET);
    if (weightClass > 0) weight = weightClass;
    const fsSelection = view.getUint16(OS2_FS_SELECTION_OFFSET);
    italic = (fsSelection & (FS_SELECTION_ITALIC | FS_SELECTION_OBLIQUE)) !== 0;
  }
  const head = tables.get("head");
  if ((weight === undefined || italic === undefined) && head && head.length >= 46) {
    const bytes = await readAt(head.offset + HEAD_MAC_STYLE_OFFSET, 2);
    const macStyle = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(0);
    weight ??= (macStyle & MAC_STYLE_BOLD) !== 0 ? BOLD_WEIGHT : DEFAULT_WEIGHT;
    italic ??= (macStyle & MAC_STYLE_ITALIC) !== 0;
  }

  return {
    filePath,
    postscriptName,
    family,
    weight: weight ?? DEFAULT_WEIGHT,
    style: italic === true ? "italic" : "normal",
  };
};

const readFontFaceFile = async (filePath: string): Promise<FontFaceRecord | undefined> => {
  const handle = await open(filePath, "r");
  try {
    return await parseFontFace(filePath, async (offset, length) => {
      const buffer = new Uint8Array(length);
      const { bytesRead } = await handle.read(buffer, 0, length, offset);
      return buffer.subarray(0, bytesRead);
    });
  } finally {
    await handle.close();
  }
};

/** Index every TrueType/OpenType face directly inside `directories`.
 * Missing directories and unreadable files are skipped. */
export const indexFontDirectories = async (
  directories: readonly string[],
): Promise<FontFaceRecord[]> => {
  const perDirectory = await Promise.all(
    directories.map(async (directory) => {
      const entries = await readdir(directory, { withFileTypes: true }).catch(() => []);
      const files = entries
        .filter(
          (entry) =>
            entry.isFile() && INDEXED_EXTENSIONS.has(path.extname(entry.name).toLowerCase()),
        )
        .map((entry) => path.join(directory, entry.name))
        .toSorted();
      const faces = await Promise.all(
        files.map((filePath) => readFontFaceFile(filePath).catch(() => undefined)),
      );
      return faces.filter((face): face is FontFaceRecord => face !== undefined);
    }),
  );
  return perDirectory.flat();
};

/** Every indexed face of the family that owns the face a PDF reports as
 * `postscriptName` (subset prefix and punctuation are ignored). Returns an
 * empty list when that face is not installed locally. */
export const findFamilyFacesForPostscriptName = (
  faces: readonly FontFaceRecord[],
  postscriptName: string,
): FontFaceRecord[] => {
  const wanted = normalizeFontName(postscriptName);
  const owner = faces.find((face) => normalizeFontName(face.postscriptName) === wanted);
  if (owner === undefined) return [];
  const family = owner.family.toLowerCase();
  return faces.filter((face) => face.family.toLowerCase() === family);
};
