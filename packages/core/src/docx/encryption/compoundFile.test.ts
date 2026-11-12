import { openCompoundFile } from "./compoundFile";

const SECTOR_BYTES = 4096;
const HEADER_DIFAT_SLOTS = 109;
const SECTOR_FREE = 0xff_ff_ff_ff;
const SECTOR_END = 0xff_ff_ff_fe;
const SECTOR_FAT = 0xff_ff_ff_fd;
const SECTOR_DIFAT = 0xff_ff_ff_fc;

type CompoundFixtureOptions = {
  /** FAT-sector ids listed in the header DIFAT and DIFAT sectors, in order. */
  fatSectorRefs: number[];
  /** Declared FAT-sector count; defaults to the number of references. */
  declaredFatSectors?: number;
};

/**
 * Minimal version-4 compound file: sector 0 is the FAT, sector 1 holds the
 * root directory entry, and any DIFAT sectors follow.
 */
const buildCompoundFile = ({
  fatSectorRefs,
  declaredFatSectors = fatSectorRefs.length,
}: CompoundFixtureOptions): Uint8Array => {
  const slotsPerDifat = SECTOR_BYTES / 4 - 1;
  const overflow = fatSectorRefs.slice(HEADER_DIFAT_SLOTS);
  const difatSectorCount = Math.ceil(overflow.length / slotsPerDifat);
  const sectorCount = 2 + difatSectorCount;
  const file = new Uint8Array(SECTOR_BYTES * (sectorCount + 1));
  const view = new DataView(file.buffer);

  file.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(0x18, 0x3e, true);
  view.setUint16(0x1a, 4, true);
  view.setUint16(0x1c, 0xff_fe, true);
  view.setUint16(0x1e, 12, true);
  view.setUint16(0x20, 6, true);
  view.setUint32(0x2c, declaredFatSectors, true);
  view.setUint32(0x30, 1, true);
  view.setUint32(0x38, 4096, true);
  view.setUint32(0x3c, SECTOR_END, true);
  view.setUint32(0x44, difatSectorCount > 0 ? 2 : SECTOR_END, true);
  view.setUint32(0x48, difatSectorCount, true);
  for (let slot = 0; slot < HEADER_DIFAT_SLOTS; slot++) {
    view.setUint32(0x4c + slot * 4, fatSectorRefs[slot] ?? SECTOR_FREE, true);
  }

  const sectorStart = (id: number): number => (id + 1) * SECTOR_BYTES;

  const fat = sectorStart(0);
  for (let entry = 0; entry < SECTOR_BYTES / 4; entry++) {
    view.setUint32(fat + entry * 4, SECTOR_FREE, true);
  }
  view.setUint32(fat, SECTOR_FAT, true);
  view.setUint32(fat + 4, SECTOR_END, true);
  for (let difat = 0; difat < difatSectorCount; difat++) {
    view.setUint32(fat + (2 + difat) * 4, SECTOR_DIFAT, true);
  }

  const root = sectorStart(1);
  const name = "Root Entry";
  for (let i = 0; i < name.length; i++) {
    view.setUint16(root + i * 2, name.charCodeAt(i), true);
  }
  view.setUint16(root + 64, (name.length + 1) * 2, true);
  file[root + 66] = 5;
  view.setUint32(root + 68, SECTOR_FREE, true);
  view.setUint32(root + 72, SECTOR_FREE, true);
  view.setUint32(root + 76, SECTOR_FREE, true);
  view.setUint32(root + 116, SECTOR_END, true);

  for (let difat = 0; difat < difatSectorCount; difat++) {
    const start = sectorStart(2 + difat);
    for (let slot = 0; slot < slotsPerDifat; slot++) {
      const ref = overflow[difat * slotsPerDifat + slot] ?? SECTOR_FREE;
      view.setUint32(start + slot * 4, ref, true);
    }
    const next = difat + 1 < difatSectorCount ? 3 + difat : SECTOR_END;
    view.setUint32(start + SECTOR_BYTES - 4, next, true);
  }

  return file;
};

describe("openCompoundFile FAT sector list", () => {
  test("opens a file whose FAT sector is listed once", () => {
    const reader = openCompoundFile(buildCompoundFile({ fatSectorRefs: [0] }));
    expect(reader.readStream("/EncryptionInfo")).toBeNull();
  });

  test("refuses a FAT sector listed more than once", () => {
    expect(() =>
      openCompoundFile(buildCompoundFile({ fatSectorRefs: [0, 0] })),
    ).toThrow(/FAT sector 0 listed more than once/);
  });

  test("refuses repeated FAT sectors listed through DIFAT sectors", () => {
    const fatSectorRefs = Array.from({ length: 4 * 1023 }, () => 0);
    expect(() =>
      openCompoundFile(buildCompoundFile({ fatSectorRefs })),
    ).toThrow(/FAT sector/);
  });

  test("refuses a FAT sector count above the sectors the file holds", () => {
    expect(() =>
      openCompoundFile(
        buildCompoundFile({
          fatSectorRefs: [0],
          declaredFatSectors: 0x10_00_00,
        }),
      ),
    ).toThrow(/FAT sector count/);
  });
});
