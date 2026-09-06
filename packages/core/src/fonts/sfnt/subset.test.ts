import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import { parseSfnt } from "./parse";
import type { SfntFont } from "./parse";
import { subsetTrueType } from "./subset";
import {
  readLocaOffsets,
  readSfntDirectory,
  SFNT_DIRECTORY_ENTRY_SIZE,
  SFNT_HEADER_SIZE,
  SFNT_TABLE,
  SFNT_VERSION,
} from "./tables";
import { TEST_FONTS_INSTALLED, TEST_FONTS_SKIP_REASON, readTestFont } from "./__tests__/testFonts";
import { decodeWoff } from "./woff";

const GLYPH_HEADER_SIZE = 10;
const COMPONENT_FLAG = {
  arg1And2AreWords: 0x0001,
  weHaveAScale: 0x0008,
  moreComponents: 0x0020,
  weHaveAnXAndYScale: 0x0040,
  weHaveATwoByTwo: 0x0080,
} as const;

const loadArimo = async (): Promise<SfntFont> => {
  const sfnt = decodeWoff(await readTestFont({ family: "arimo" }));
  if (sfnt.isErr()) {
    throw sfnt.error;
  }
  const font = parseSfnt(sfnt.value);
  if (font.isErr()) {
    throw font.error;
  }
  return font.value;
};

/** Independent glyph reader, so the tests check the bytes, not the subsetter. */
const glyphDataFor = (font: SfntFont, glyphId: number): Uint8Array | undefined => {
  const directory = readSfntDirectory(font.bytes);
  if (directory.isErr()) {
    return undefined;
  }

  const { view, tables } = directory.value;
  const head = tables.get(SFNT_TABLE.head);
  const loca = tables.get(SFNT_TABLE.loca);
  const glyf = tables.get(SFNT_TABLE.glyf);
  if (!head || !loca || !glyf) {
    return undefined;
  }

  const offsets = readLocaOffsets({
    view,
    loca,
    numGlyphs: font.numGlyphs,
    longFormat: view.getInt16(head.offset + 50) === 1,
  });
  const start = offsets?.[glyphId];
  const end = offsets?.[glyphId + 1];
  if (start === undefined || end === undefined || end <= start) {
    return undefined;
  }
  return font.bytes.subarray(glyf.offset + start, glyf.offset + end);
};

/** Component glyph ids of a composite glyph, or [] for a simple one. */
const componentGlyphIds = (glyph: Uint8Array): readonly number[] => {
  const view = new DataView(glyph.buffer, glyph.byteOffset, glyph.byteLength);
  if (view.getInt16(0) >= 0) {
    return [];
  }

  const ids: number[] = [];
  let at = GLYPH_HEADER_SIZE;
  while (at + 4 <= glyph.byteLength) {
    const flags = view.getUint16(at);
    ids.push(view.getUint16(at + 2));
    at += 4;
    at += (flags & COMPONENT_FLAG.arg1And2AreWords) === 0 ? 2 : 4;
    if ((flags & COMPONENT_FLAG.weHaveAScale) !== 0) {
      at += 2;
    } else if ((flags & COMPONENT_FLAG.weHaveAnXAndYScale) !== 0) {
      at += 4;
    } else if ((flags & COMPONENT_FLAG.weHaveATwoByTwo) !== 0) {
      at += 8;
    }
    if ((flags & COMPONENT_FLAG.moreComponents) === 0) {
      break;
    }
  }
  return ids;
};

/** The first Latin-1 accented character in the font whose glyph is composite. */
const findCompositeGlyph = (font: SfntFont): number | undefined => {
  for (let codePoint = 0x00c0; codePoint <= 0x00ff; codePoint++) {
    const glyphId = font.glyphIdFor(codePoint);
    if (glyphId === 0) {
      continue;
    }
    const glyph = glyphDataFor(font, glyphId);
    if (glyph && glyph.byteLength >= GLYPH_HEADER_SIZE && componentGlyphIds(glyph).length > 0) {
      return glyphId;
    }
  }
  return undefined;
};

const glyphIdsFor = (font: SfntFont, text: string): Set<number> => {
  const ids = new Set<number>();
  for (const character of text) {
    ids.add(font.glyphIdFor(character.codePointAt(0) ?? 0));
  }
  return ids;
};

/** Minimal sfnt builder, used to stand up fonts the fixtures do not provide. */
const buildSfnt = (version: number, tables: readonly { tag: string; data: Uint8Array }[]) => {
  const sorted = [...tables].sort((left, right) => (left.tag < right.tag ? -1 : 1));
  const directorySize = sorted.length * SFNT_DIRECTORY_ENTRY_SIZE;
  const bodySize = sorted.reduce(
    (total, table) => total + Math.ceil(table.data.byteLength / 4) * 4,
    0,
  );
  const output = new Uint8Array(SFNT_HEADER_SIZE + directorySize + bodySize);
  const view = new DataView(output.buffer);
  view.setUint32(0, version);
  view.setUint16(4, sorted.length);

  let entry = SFNT_HEADER_SIZE;
  let dataOffset = SFNT_HEADER_SIZE + directorySize;
  for (const table of sorted) {
    for (let index = 0; index < table.tag.length; index++) {
      view.setUint8(entry + index, table.tag.charCodeAt(index));
    }
    view.setUint32(entry + 8, dataOffset);
    view.setUint32(entry + 12, table.data.byteLength);
    output.set(table.data, dataOffset);
    dataOffset += Math.ceil(table.data.byteLength / 4) * 4;
    entry += SFNT_DIRECTORY_ENTRY_SIZE;
  }
  return output;
};

const HEAD_INDEX_TO_LOC_FORMAT = 50;
const HHEA_NUMBER_OF_HMETRICS = 34;
const MAXP_NUM_GLYPHS = 4;
const MAX_SHORT_LOCA_GLYF_LENGTH = 0x1fffe;

/** Enough filler glyphs to push `glyf` past what short loca can address. */
const FILLER_GLYPH_COUNT = 96;
const FILLER_GLYPH_SIZE = 2048;
const FILLER_ADVANCE = 600;

/**
 * A synthetic TrueType font: the donor's head/hhea/maxp with a `glyf` of
 * fixed-size filler glyphs, large enough that a subset must use long loca.
 * The glyph bodies are never interpreted, only copied.
 */
const buildFillerFont = (donor: SfntFont): SfntFont => {
  const directory = readSfntDirectory(donor.bytes);
  if (directory.isErr()) {
    panic(directory.error.message);
  }
  const copy = (tag: string) => {
    const record = directory.value.tables.get(tag);
    if (!record) {
      panic(`donor font has no '${tag}'`);
    }
    return donor.bytes.slice(record.offset, record.offset + record.length);
  };

  const head = copy(SFNT_TABLE.head);
  new DataView(head.buffer).setInt16(HEAD_INDEX_TO_LOC_FORMAT, 1);
  const hhea = copy(SFNT_TABLE.hhea);
  new DataView(hhea.buffer).setUint16(HHEA_NUMBER_OF_HMETRICS, FILLER_GLYPH_COUNT);
  const maxp = copy(SFNT_TABLE.maxp);
  new DataView(maxp.buffer).setUint16(MAXP_NUM_GLYPHS, FILLER_GLYPH_COUNT);

  const hmtx = new Uint8Array(FILLER_GLYPH_COUNT * 4);
  const hmtxView = new DataView(hmtx.buffer);
  const glyf = new Uint8Array(FILLER_GLYPH_COUNT * FILLER_GLYPH_SIZE);
  const glyfView = new DataView(glyf.buffer);
  const loca = new Uint8Array((FILLER_GLYPH_COUNT + 1) * 4);
  const locaView = new DataView(loca.buffer);
  for (let glyphId = 0; glyphId < FILLER_GLYPH_COUNT; glyphId++) {
    hmtxView.setUint16(glyphId * 4, FILLER_ADVANCE + glyphId);
    hmtxView.setInt16(glyphId * 4 + 2, glyphId);
    locaView.setUint32(glyphId * 4, glyphId * FILLER_GLYPH_SIZE);

    const at = glyphId * FILLER_GLYPH_SIZE;
    glyfView.setInt16(at, 1);
    glyfView.setInt16(at + 2, 0);
    glyfView.setInt16(at + 4, -50);
    glyfView.setInt16(at + 6, 400);
    glyfView.setInt16(at + 8, 700 + glyphId);
  }
  locaView.setUint32(FILLER_GLYPH_COUNT * 4, FILLER_GLYPH_COUNT * FILLER_GLYPH_SIZE);

  const parsed = parseSfnt(
    buildSfnt(SFNT_VERSION.trueType, [
      { tag: SFNT_TABLE.head, data: head },
      { tag: SFNT_TABLE.hhea, data: hhea },
      { tag: SFNT_TABLE.maxp, data: maxp },
      { tag: SFNT_TABLE.hmtx, data: hmtx },
      { tag: SFNT_TABLE.loca, data: loca },
      { tag: SFNT_TABLE.glyf, data: glyf },
    ]),
  );
  if (parsed.isErr()) {
    panic(parsed.error.message);
  }
  return parsed.value;
};

/** The donor's own head/hhea/maxp/hmtx, re-wrapped as an `OTTO` CFF font. */
const buildCffLookalike = (font: SfntFont): Uint8Array => {
  const directory = readSfntDirectory(font.bytes);
  if (directory.isErr()) {
    throw directory.error;
  }
  const copy = (tag: string) => {
    const record = directory.value.tables.get(tag);
    if (!record) {
      panic(`donor font has no '${tag}'`);
    }
    return font.bytes.slice(record.offset, record.offset + record.length);
  };

  return buildSfnt(SFNT_VERSION.openTypeCff, [
    { tag: SFNT_TABLE.head, data: copy(SFNT_TABLE.head) },
    { tag: SFNT_TABLE.hhea, data: copy(SFNT_TABLE.hhea) },
    { tag: SFNT_TABLE.maxp, data: copy(SFNT_TABLE.maxp) },
    { tag: SFNT_TABLE.hmtx, data: copy(SFNT_TABLE.hmtx) },
    { tag: SFNT_TABLE.cff, data: new Uint8Array(16) },
  ]);
};

describe(`subsetTrueType (${TEST_FONTS_SKIP_REASON})`, () => {
  test.skipIf(!TEST_FONTS_INSTALLED)("produces a font that re-parses", async () => {
    const font = await loadArimo();
    const requested = glyphIdsFor(font, "Hello, world!");
    const subset = subsetTrueType(font, requested);

    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    const reparsed = parseSfnt(subset.value.bytes);
    expect(reparsed.isErr()).toBe(false);
    if (reparsed.isErr()) {
      return;
    }

    expect(reparsed.value.numGlyphs).toBe(subset.value.glyphIdMap.size);
    expect(reparsed.value.unitsPerEm).toBe(font.unitsPerEm);
    expect(reparsed.value.ascender).toBe(font.ascender);
    expect(reparsed.value.bbox).toEqual(font.bbox);
    expect(subset.value.bytes.byteLength).toBeLessThan(font.bytes.byteLength);
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("keeps .notdef at glyph 0", async () => {
    const font = await loadArimo();
    const subset = subsetTrueType(font, glyphIdsFor(font, "xyz"));
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    expect(subset.value.glyphIdMap.get(0)).toBe(0);
    // New ids are a dense 0..n-1 range.
    expect([...subset.value.glyphIdMap.values()].sort((left, right) => left - right)).toEqual(
      [...subset.value.glyphIdMap.values()].map((_, index) => index),
    );
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("preserves the requested advances", async () => {
    const font = await loadArimo();
    const subset = subsetTrueType(font, glyphIdsFor(font, "The quick brown fox, 1234!"));
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    const reparsed = parseSfnt(subset.value.bytes);
    expect(reparsed.isErr()).toBe(false);
    if (reparsed.isErr()) {
      return;
    }

    for (const [oldId, newId] of subset.value.glyphIdMap) {
      expect(reparsed.value.advanceWidthFor(newId)).toBe(font.advanceWidthFor(oldId));
    }
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("keeps the glyph outlines byte for byte", async () => {
    const font = await loadArimo();
    const subset = subsetTrueType(font, glyphIdsFor(font, "Wg"));
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    const reparsed = parseSfnt(subset.value.bytes);
    expect(reparsed.isErr()).toBe(false);
    if (reparsed.isErr()) {
      return;
    }

    for (const [oldId, newId] of subset.value.glyphIdMap) {
      const original = glyphDataFor(font, oldId);
      const copied = glyphDataFor(reparsed.value, newId);
      if (!original) {
        continue;
      }
      expect(copied).toBeDefined();
      // Composite components are renumbered, so only simple glyphs compare
      // byte for byte; the copy is padded to four bytes.
      if (componentGlyphIds(original).length === 0 && copied) {
        expect(
          Buffer.from(copied.subarray(0, original.byteLength)).equals(Buffer.from(original)),
        ).toBe(true);
      }
      expect(reparsed.value.glyphBoundsFor(newId)).toEqual(font.glyphBoundsFor(oldId));
    }
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("pulls composite components into the subset", async () => {
    const font = await loadArimo();
    const compositeId = findCompositeGlyph(font);
    expect(compositeId).toBeDefined();
    if (compositeId === undefined) {
      return;
    }

    const original = glyphDataFor(font, compositeId);
    expect(original).toBeDefined();
    if (!original) {
      return;
    }
    const components = componentGlyphIds(original);
    expect(components.length).toBeGreaterThan(0);

    const subset = subsetTrueType(font, new Set([compositeId]));
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    for (const component of components) {
      expect(subset.value.glyphIdMap.has(component)).toBe(true);
    }

    const reparsed = parseSfnt(subset.value.bytes);
    expect(reparsed.isErr()).toBe(false);
    if (reparsed.isErr()) {
      return;
    }

    const newId = subset.value.glyphIdMap.get(compositeId);
    expect(newId).toBeDefined();
    if (newId === undefined) {
      return;
    }

    const copied = glyphDataFor(reparsed.value, newId);
    expect(copied).toBeDefined();
    if (!copied) {
      return;
    }

    // The component indices must have been rewritten to the new numbering.
    const copiedComponents = componentGlyphIds(copied);
    expect(copiedComponents.length).toBe(components.length);
    for (const [index, component] of components.entries()) {
      expect(copiedComponents[index]).toBe(subset.value.glyphIdMap.get(component));
      expect(copiedComponents[index]).toBeLessThan(reparsed.value.numGlyphs);
    }
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("is byte-identical across runs and set orders", async () => {
    const font = await loadArimo();
    const ascending = glyphIdsFor(font, "abcdefghij");
    const shuffled = new Set([...ascending].reverse());

    const first = subsetTrueType(font, ascending);
    const second = subsetTrueType(font, ascending);
    const reordered = subsetTrueType(font, shuffled);
    expect(first.isErr()).toBe(false);
    expect(second.isErr()).toBe(false);
    expect(reordered.isErr()).toBe(false);
    if (first.isErr() || second.isErr() || reordered.isErr()) {
      return;
    }

    expect(Buffer.from(first.value.bytes).equals(Buffer.from(second.value.bytes))).toBe(true);
    expect(Buffer.from(first.value.bytes).equals(Buffer.from(reordered.value.bytes))).toBe(true);
    expect([...first.value.glyphIdMap]).toEqual([...reordered.value.glyphIdMap]);
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("writes a checkSumAdjustment the file sums to", async () => {
    const font = await loadArimo();
    const subset = subsetTrueType(font, glyphIdsFor(font, "checksum"));
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    const bytes = subset.value.bytes;
    expect(bytes.byteLength % 4).toBe(0);

    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let sum = 0;
    for (let at = 0; at < bytes.byteLength; at += 4) {
      sum = (sum + view.getUint32(at)) % 0x100000000;
    }
    expect(sum).toBe(0xb1b0afba);
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("subsets an empty request down to .notdef", async () => {
    const font = await loadArimo();
    const subset = subsetTrueType(font, new Set());
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    expect(subset.value.glyphIdMap.size).toBe(1);
    const reparsed = parseSfnt(subset.value.bytes);
    expect(reparsed.isErr()).toBe(false);
    if (!reparsed.isErr()) {
      expect(reparsed.value.numGlyphs).toBe(1);
    }
  });

  test.skipIf(!TEST_FONTS_INSTALLED)(
    "copies every glyph when everything is requested",
    async () => {
      const font = await loadArimo();
      const everyGlyph = new Set(Array.from({ length: font.numGlyphs }, (_, index) => index));
      const subset = subsetTrueType(font, everyGlyph);
      expect(subset.isErr()).toBe(false);
      if (subset.isErr()) {
        return;
      }

      const reparsed = parseSfnt(subset.value.bytes);
      expect(reparsed.isErr()).toBe(false);
      if (reparsed.isErr()) {
        return;
      }
      expect(reparsed.value.numGlyphs).toBe(font.numGlyphs);
      // A full subset renumbers nothing, so every glyph keeps its own id.
      for (const [oldId, newId] of subset.value.glyphIdMap) {
        expect(newId).toBe(oldId);
        expect(reparsed.value.glyphBoundsFor(newId)).toEqual(font.glyphBoundsFor(oldId));
      }
    },
  );

  test.skipIf(!TEST_FONTS_INSTALLED)("switches to long loca for a large glyf", async () => {
    // The Latin fixture faces are far too small to overrun short loca, so the
    // branch that would silently truncate a large font gets its own font.
    const font = buildFillerFont(await loadArimo());
    const subset = subsetTrueType(
      font,
      new Set(Array.from({ length: font.numGlyphs }, (_, index) => index)),
    );
    expect(subset.isErr()).toBe(false);
    if (subset.isErr()) {
      return;
    }

    const directory = readSfntDirectory(subset.value.bytes);
    expect(directory.isErr()).toBe(false);
    if (directory.isErr()) {
      return;
    }

    const { view, tables } = directory.value;
    const head = tables.get(SFNT_TABLE.head);
    const glyf = tables.get(SFNT_TABLE.glyf);
    const loca = tables.get(SFNT_TABLE.loca);
    expect(head).toBeDefined();
    expect(glyf).toBeDefined();
    expect(loca).toBeDefined();
    if (!head || !glyf || !loca) {
      return;
    }

    expect(glyf.length).toBeGreaterThan(MAX_SHORT_LOCA_GLYF_LENGTH);
    expect(view.getInt16(head.offset + HEAD_INDEX_TO_LOC_FORMAT)).toBe(1);
    expect(loca.length).toBe((font.numGlyphs + 1) * 4);

    const reparsed = parseSfnt(subset.value.bytes);
    expect(reparsed.isErr()).toBe(false);
    if (reparsed.isErr()) {
      return;
    }
    expect(reparsed.value.numGlyphs).toBe(font.numGlyphs);
    for (let glyphId = 0; glyphId < font.numGlyphs; glyphId++) {
      expect(reparsed.value.advanceWidthFor(glyphId)).toBe(font.advanceWidthFor(glyphId));
      expect(reparsed.value.glyphBoundsFor(glyphId)).toEqual(font.glyphBoundsFor(glyphId));
    }
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("rejects a glyph id outside the font", async () => {
    const font = await loadArimo();
    const subset = subsetTrueType(font, new Set([font.numGlyphs]));
    expect(subset.isErr()).toBe(true);
    if (subset.isErr()) {
      expect(subset.error.message).toContain("outside");
    }
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("refuses a CFF font", async () => {
    const donor = await loadArimo();
    const cff = parseSfnt(buildCffLookalike(donor));
    expect(cff.isErr()).toBe(false);
    if (cff.isErr()) {
      return;
    }

    expect(cff.value.isCff).toBe(true);
    const subset = subsetTrueType(cff.value, new Set([1]));
    expect(subset.isErr()).toBe(true);
    if (subset.isErr()) {
      expect(subset.error.message).toContain("CFF");
    }
  });
});
