/**
 * Read-only MS-CFB compound file access for encrypted OOXML packages.
 *
 * Encrypted `.docx` files store `EncryptionInfo` and `EncryptedPackage` streams
 * at the compound-file root. This reader implements only what MS-OFFCRYPTO needs.
 *
 * @see https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
 */

import { joinBytes } from "./cryptoBytes";

const HEADER_BYTES = 512;
const DIRECTORY_RECORD_BYTES = 128;
const HEADER_DIFAT_SLOTS = 109;

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;
const LITTLE_ENDIAN_MARK = 0xff_fe;
const SECTOR_FREE = 0xff_ff_ff_ff;
const SECTOR_END = 0xff_ff_ff_fe;
const SECTOR_DIFAT = 0xff_ff_ff_fc;
const DIR_NONE = 0xff_ff_ff_ff;

const ENTRY_KIND = {
  STORAGE: 1,
  STREAM: 2,
  ROOT: 5,
} as const;

type CompoundHeader = {
  sectorBytes: number;
  miniSectorBytes: number;
  fatSectorCount: number;
  directoryStartSector: number;
  miniCutoff: number;
  miniFatStart: number;
  miniFatSectorCount: number;
  difatStart: number;
  difatSectorCount: number;
  headerDifat: number[];
};

type DirectoryRecord = {
  name: string;
  kind: number;
  leftId: number;
  rightId: number;
  childId: number;
  firstSector: number;
  byteLength: number;
};

const viewOf = (bytes: Uint8Array): DataView =>
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

const u16 = (view: DataView, offset: number): number => view.getUint16(offset, true);
const u32 = (view: DataView, offset: number): number => view.getUint32(offset, true);
const u64 = (view: DataView, offset: number): bigint => {
  const low = BigInt(view.getUint32(offset, true));
  const high = BigInt(view.getUint32(offset + 4, true));
  return (high << 32n) | low;
};

const decodeName = (nameBytes: Uint8Array): string => {
  const decoded = new TextDecoder("utf-16le").decode(nameBytes);
  let end = decoded.length;
  while (end > 0 && decoded.charCodeAt(end - 1) === 0) {
    end--;
  }
  return decoded.slice(0, end);
};

const sectorOffset = (sectorId: number, sectorBytes: number): number =>
  (sectorId + 1) * sectorBytes;

const readHeader = (file: Uint8Array): CompoundHeader => {
  if (file.length < HEADER_BYTES) {
    throw new Error(`Compound file header truncated (${file.length} bytes)`);
  }
  for (let i = 0; i < SIGNATURE.length; i++) {
    if (file[i] !== SIGNATURE[i]) {
      throw new Error("Invalid compound file signature");
    }
  }

  const view = viewOf(file);
  const major = u16(view, 0x1a);
  const endian = u16(view, 0x1c);
  const sectorShift = u16(view, 0x1e);
  const miniShift = u16(view, 0x20);
  const sectorBytes = 1 << sectorShift;
  const miniSectorBytes = 1 << miniShift;

  if (endian !== LITTLE_ENDIAN_MARK) {
    throw new Error(`Unsupported compound file endianness: 0x${endian.toString(16)}`);
  }
  if (!((major === 3 && sectorBytes === 512) || (major === 4 && sectorBytes === 4096))) {
    throw new Error(`Unsupported compound file version ${major} (sector size ${sectorBytes})`);
  }

  const headerDifat: number[] = [];
  for (let slot = 0; slot < HEADER_DIFAT_SLOTS; slot++) {
    const id = u32(view, 0x4c + slot * 4);
    if (id !== SECTOR_FREE) {
      headerDifat.push(id);
    }
  }

  return {
    sectorBytes,
    miniSectorBytes,
    fatSectorCount: u32(view, 0x2c),
    directoryStartSector: u32(view, 0x30),
    miniCutoff: u32(view, 0x38),
    miniFatStart: u32(view, 0x3c),
    miniFatSectorCount: u32(view, 0x40),
    difatStart: u32(view, 0x44),
    difatSectorCount: u32(view, 0x48),
    headerDifat,
  };
};

const loadSector = (file: Uint8Array, header: CompoundHeader, sectorId: number): Uint8Array => {
  if (sectorId >= SECTOR_DIFAT) {
    throw new Error(`Invalid sector id 0x${sectorId.toString(16)}`);
  }
  const offset = sectorOffset(sectorId, header.sectorBytes);
  if (offset + header.sectorBytes > file.length) {
    throw new Error(`Sector ${sectorId} extends past end of file`);
  }
  return file.subarray(offset, offset + header.sectorBytes);
};

const loadFatSectorIds = (file: Uint8Array, header: CompoundHeader): number[] => {
  const ids = [...header.headerDifat];
  // A hostile `difatSectorCount` (up to 2^32-1) would otherwise force this
  // loop to run far more times than the file could possibly contain DIFAT
  // sectors; a chain that revisits a sector could spin for its full
  // declared length even at a capped count. Bound the loop by the sectors
  // the file can actually hold and guard against a cycle, mirroring
  // `followSectorChain`.
  const maxDifatSectors = Math.ceil(file.length / header.sectorBytes);
  const stepLimit = Math.min(header.difatSectorCount, maxDifatSectors);
  const seen = new Set<number>();
  let chain = header.difatStart;
  for (let step = 0; step < stepLimit; step++) {
    if (chain === SECTOR_END || chain === SECTOR_FREE) {
      break;
    }
    if (seen.has(chain)) {
      throw new Error(`DIFAT chain loop at sector ${chain}`);
    }
    seen.add(chain);
    const sector = loadSector(file, header, chain);
    const view = viewOf(sector);
    const slotsPerSector = header.sectorBytes / 4 - 1;
    for (let slot = 0; slot < slotsPerSector; slot++) {
      const id = u32(view, slot * 4);
      if (id !== SECTOR_FREE) {
        ids.push(id);
      }
    }
    chain = u32(view, header.sectorBytes - 4);
  }
  if (ids.length < header.fatSectorCount) {
    throw new Error("DIFAT chain shorter than declared FAT sector count");
  }
  return ids.slice(0, header.fatSectorCount);
};

const loadFat = (file: Uint8Array, header: CompoundHeader): number[] => {
  const table: number[] = [];
  for (const fatSectorId of loadFatSectorIds(file, header)) {
    const sector = loadSector(file, header, fatSectorId);
    const view = viewOf(sector);
    const entries = header.sectorBytes / 4;
    for (let i = 0; i < entries; i++) {
      table.push(u32(view, i * 4));
    }
  }
  return table;
};

const followSectorChain = (
  file: Uint8Array,
  header: CompoundHeader,
  fat: number[],
  start: number,
): Uint8Array => {
  if (start === SECTOR_END || start === SECTOR_FREE) {
    return new Uint8Array(0);
  }

  const parts: Uint8Array[] = [];
  const seen = new Set<number>();
  let sectorId = start;

  while (sectorId !== SECTOR_END) {
    if (seen.has(sectorId)) {
      throw new Error(`Sector chain loop at sector ${sectorId}`);
    }
    if (sectorId >= fat.length) {
      throw new Error(`Sector ${sectorId} outside FAT`);
    }
    seen.add(sectorId);
    parts.push(loadSector(file, header, sectorId));
    const next = fat[sectorId];
    if (next === undefined) {
      throw new Error(`FAT missing link after sector ${sectorId}`);
    }
    sectorId = next;
  }

  return joinBytes(parts);
};

const loadDirectory = (
  file: Uint8Array,
  header: CompoundHeader,
  fat: number[],
): DirectoryRecord[] => {
  const raw = followSectorChain(file, header, fat, header.directoryStartSector);
  const records: DirectoryRecord[] = [];

  for (let pos = 0; pos + DIRECTORY_RECORD_BYTES <= raw.length; pos += DIRECTORY_RECORD_BYTES) {
    const slice = raw.subarray(pos, pos + DIRECTORY_RECORD_BYTES);
    const view = viewOf(slice);
    const nameChars = Math.max(0, Math.min(u16(view, 64), 64) - 2);
    records.push({
      name: decodeName(slice.subarray(0, nameChars)),
      kind: slice[66] ?? 0,
      leftId: u32(view, 68),
      rightId: u32(view, 72),
      childId: u32(view, 76),
      firstSector: u32(view, 116),
      byteLength: (() => {
        const size = u64(view, 120);
        if (size > BigInt(Number.MAX_SAFE_INTEGER)) {
          throw new Error(`Stream too large: ${size.toString()}`);
        }
        return Number(size);
      })(),
    });
  }

  return records;
};

/** Breadth-first sibling-tree walk (MS-CFB red-black sibling links). */
const walkSiblingTree = (
  records: DirectoryRecord[],
  rootId: number,
  visit: (record: DirectoryRecord) => void,
): void => {
  const queue: number[] = [];
  const enqueueSubtree = (id: number): void => {
    if (id === DIR_NONE || id >= records.length) {
      return;
    }
    const stack = [id];
    const queued = new Set<number>();
    while (stack.length > 0) {
      const current = stack.pop()!;
      if (queued.has(current)) {
        throw new Error(`Directory sibling loop at entry ${current}`);
      }
      queued.add(current);
      queue.push(current);
      const node = records[current];
      if (!node) {
        continue;
      }
      if (node.leftId !== DIR_NONE) {
        stack.push(node.leftId);
      }
      if (node.rightId !== DIR_NONE) {
        stack.push(node.rightId);
      }
    }
  };
  enqueueSubtree(rootId);
  for (const id of queue) {
    const record = records[id];
    if (record) {
      visit(record);
    }
  }
};

const resolvePath = (records: DirectoryRecord[], path: string): DirectoryRecord | null => {
  const parts = path.split("/").filter(Boolean);
  const root = records[0];
  if (!root) {
    return null;
  }

  let current = root;
  for (const part of parts) {
    if (current.kind !== ENTRY_KIND.ROOT && current.kind !== ENTRY_KIND.STORAGE) {
      return null;
    }
    let match: DirectoryRecord | null = null;
    walkSiblingTree(records, current.childId, (candidate) => {
      if (candidate.name.toLowerCase() === part.toLowerCase()) {
        match = candidate;
      }
    });
    if (!match) {
      return null;
    }
    current = match;
  }
  return current;
};

const loadMiniFat = (file: Uint8Array, header: CompoundHeader, fat: number[]): number[] => {
  if (header.miniFatSectorCount === 0 || header.miniFatStart === SECTOR_END) {
    return [];
  }
  const bytes = followSectorChain(file, header, fat, header.miniFatStart);
  const view = viewOf(bytes);
  const count = Math.floor(bytes.length / 4);
  const entries: number[] = [];
  for (let i = 0; i < count; i++) {
    entries.push(u32(view, i * 4));
  }
  return entries;
};

const followMiniChain = (
  container: Uint8Array,
  miniFat: number[],
  miniSectorBytes: number,
  start: number,
): Uint8Array => {
  if (start === SECTOR_END || start === SECTOR_FREE) {
    return new Uint8Array(0);
  }
  const parts: Uint8Array[] = [];
  const seen = new Set<number>();
  let sectorId = start;
  while (sectorId !== SECTOR_END) {
    if (seen.has(sectorId)) {
      throw new Error(`Mini sector chain loop at ${sectorId}`);
    }
    if (sectorId >= miniFat.length) {
      throw new Error(`Mini sector ${sectorId} outside MiniFAT`);
    }
    seen.add(sectorId);
    const offset = sectorId * miniSectorBytes;
    if (offset + miniSectorBytes > container.length) {
      throw new Error(`Mini sector ${sectorId} extends past mini stream`);
    }
    parts.push(container.subarray(offset, offset + miniSectorBytes));
    const next = miniFat[sectorId];
    if (next === undefined) {
      throw new Error(`MiniFAT missing link after mini sector ${sectorId}`);
    }
    sectorId = next;
  }
  return joinBytes(parts);
};

export type CompoundFileReader = {
  readStream: (absolutePath: string) => Uint8Array | null;
};

export const openCompoundFile = (data: ArrayBuffer | Uint8Array): CompoundFileReader => {
  const file = data instanceof Uint8Array ? data : new Uint8Array(data);
  const header = readHeader(file);
  const fat = loadFat(file, header);
  const directory = loadDirectory(file, header, fat);
  const root = directory[0];
  if (!root || root.kind !== ENTRY_KIND.ROOT) {
    throw new Error("Compound file root entry missing");
  }

  const miniFat = loadMiniFat(file, header, fat);
  const miniContainer =
    root.byteLength > 0 && root.firstSector !== SECTOR_END
      ? followSectorChain(file, header, fat, root.firstSector).subarray(0, root.byteLength)
      : new Uint8Array(0);

  return {
    readStream(absolutePath: string): Uint8Array | null {
      const entry = resolvePath(directory, absolutePath);
      if (!entry || entry.kind !== ENTRY_KIND.STREAM) {
        return null;
      }
      if (entry.byteLength < header.miniCutoff) {
        return followMiniChain(
          miniContainer,
          miniFat,
          header.miniSectorBytes,
          entry.firstSector,
        ).subarray(0, entry.byteLength);
      }
      return followSectorChain(file, header, fat, entry.firstSector).subarray(0, entry.byteLength);
    },
  };
};

export type EncryptedPackageStreams = {
  encryptionInfo: Uint8Array;
  encryptedPackage: Uint8Array;
};

export const readEncryptedPackageStreams = (
  data: ArrayBuffer | Uint8Array,
): EncryptedPackageStreams => {
  const file = data instanceof Uint8Array ? data : new Uint8Array(data);
  let reader: CompoundFileReader;
  try {
    reader = openCompoundFile(file);
  } catch (cause) {
    throw new Error("Failed to parse encrypted OOXML compound file", { cause });
  }

  const encryptionInfo = reader.readStream("/EncryptionInfo");
  const encryptedPackage = reader.readStream("/EncryptedPackage");
  if (!encryptionInfo) {
    throw new Error("Compound file missing EncryptionInfo stream");
  }
  if (!encryptedPackage) {
    throw new Error("Compound file missing EncryptedPackage stream");
  }
  return { encryptionInfo, encryptedPackage };
};
