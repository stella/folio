/**
 * TrueType (`glyf`) subsetting.
 *
 * The output is what a PDF CIDFontType2 needs and nothing else: glyph
 * outlines, metrics, and the hinting tables that go with them. `cmap`, `name`
 * and `post` are dropped, because a CIDFontType2 addresses glyphs by index
 * through its own CIDToGIDMap and never consults the font's own character
 * mapping.
 *
 * Determinism contract: `subsetTrueType(font, glyphIds)` is a pure function of
 * its arguments and produces byte-identical output across runs. Glyph ids are
 * iterated in ascending numeric order and tables are written in tag order, so
 * nothing depends on a Set's or Map's insertion order.
 */

import { panic, Result, TaggedError } from "better-result";

import type { SfntFont } from "./parse";
import {
  fitsInView,
  readLocaOffsets,
  readSfntDirectory,
  SFNT_DIRECTORY_ENTRY_SIZE,
  SFNT_HEADER_SIZE,
  SFNT_TABLE,
  SFNT_TABLE_ALIGNMENT,
  SFNT_VERSION,
  type SfntDirectory,
  type SfntTableRecord,
} from "./tables";

export class SubsetError extends TaggedError("SubsetError")<{ message: string }> {}

export type TrueTypeSubset = {
  /** Subset sfnt bytes. */
  readonly bytes: Uint8Array;
  /** Old glyph id -> new glyph id. Always maps 0 -> 0. */
  readonly glyphIdMap: ReadonlyMap<number, number>;
};

const fail = (message: string) => Result.err(new SubsetError({ message }));

/** .notdef is glyph 0 in every font and is always kept. */
const NOTDEF_GLYPH_ID = 0;

const HEAD = { checkSumAdjustment: 8, indexToLocFormat: 50, size: 54 } as const;
const HHEA = { numberOfHMetrics: 34, size: 36 } as const;
const MAXP = { numGlyphs: 4, size: 6 } as const;

/** One hmtx `longHorMetric`: advanceWidth (uint16) then lsb (int16). */
const HMTX_METRIC_SIZE = 4;

/** Glyph header: numberOfContours, then the four bounding-box int16s. */
const GLYPH_HEADER_SIZE = 10;
const COMPOSITE_CONTOUR_COUNT = -1;

/** Flags of a composite glyph component record. */
const COMPONENT_FLAG = {
  arg1And2AreWords: 0x0001,
  weHaveAScale: 0x0008,
  moreComponents: 0x0020,
  weHaveAnXAndYScale: 0x0040,
  weHaveATwoByTwo: 0x0080,
} as const;

/**
 * Short `loca` stores offsets halved in a uint16, so it can only address
 * 0x1FFFE bytes of `glyf`. Beyond that the long format is mandatory.
 */
const MAX_SHORT_LOCA_GLYF_LENGTH = 0x1fffe;

/** `head.checkSumAdjustment` is this constant minus the whole-file checksum. */
const CHECKSUM_ADJUSTMENT_MAGIC = 0xb1b0afba;

const UINT32_MODULUS = 0x100000000;

/** Tables copied byte for byte when the source font has them. */
const VERBATIM_TABLES = [SFNT_TABLE.cvt, SFNT_TABLE.fpgm, SFNT_TABLE.prep] as const;

const alignUp = (value: number): number =>
  Math.ceil(value / SFNT_TABLE_ALIGNMENT) * SFNT_TABLE_ALIGNMENT;

/**
 * The offset of every component's glyph index inside one composite glyph, in
 * the order the records appear. Returned separately from the referenced ids so
 * the closure walk and the renumbering pass share one parser.
 */
type CompositeComponents = {
  readonly glyphIndexOffsets: readonly number[];
  readonly glyphIds: readonly number[];
};

const readCompositeComponents = (
  view: DataView,
  start: number,
  end: number,
): CompositeComponents | undefined => {
  const glyphIndexOffsets: number[] = [];
  const glyphIds: number[] = [];
  let at = start + GLYPH_HEADER_SIZE;
  for (;;) {
    if (at + 4 > end || !fitsInView(view, at, 4)) {
      return undefined;
    }

    const flags = view.getUint16(at);
    glyphIndexOffsets.push(at + 2);
    glyphIds.push(view.getUint16(at + 2));
    at += 4;

    at += (flags & COMPONENT_FLAG.arg1And2AreWords) === 0 ? 2 : 4;
    if ((flags & COMPONENT_FLAG.weHaveAScale) !== 0) {
      at += 2;
    } else if ((flags & COMPONENT_FLAG.weHaveAnXAndYScale) !== 0) {
      at += 4;
    } else if ((flags & COMPONENT_FLAG.weHaveATwoByTwo) !== 0) {
      at += 8;
    }
    if (at > end) {
      return undefined;
    }
    if ((flags & COMPONENT_FLAG.moreComponents) === 0) {
      return { glyphIndexOffsets, glyphIds };
    }
  }
};

type GlyphSource = {
  readonly view: DataView;
  readonly glyf: SfntTableRecord;
  readonly offsets: readonly number[];
  readonly numGlyphs: number;
};

/** Byte range of one glyph inside `glyf`, or `undefined` for an empty glyph. */
const glyphRange = (
  { glyf, offsets }: GlyphSource,
  glyphId: number,
): { readonly start: number; readonly end: number } | undefined => {
  const start = offsets[glyphId];
  const end = offsets[glyphId + 1];
  if (start === undefined || end === undefined || end <= start || end > glyf.length) {
    return undefined;
  }
  return { start: glyf.offset + start, end: glyf.offset + end };
};

/**
 * The requested glyphs plus every glyph reachable through composite component
 * references, plus .notdef, sorted ascending.
 */
const closeOverComponents = (
  source: GlyphSource,
  requested: ReadonlySet<number>,
): Result<readonly number[], SubsetError> => {
  const selected = new Set<number>([NOTDEF_GLYPH_ID]);
  const queue: number[] = [NOTDEF_GLYPH_ID];
  for (const glyphId of [...requested].sort((left, right) => left - right)) {
    if (!Number.isInteger(glyphId) || glyphId < 0 || glyphId >= source.numGlyphs) {
      return fail(`glyph id ${glyphId} is outside 0..${source.numGlyphs - 1}`);
    }
    if (!selected.has(glyphId)) {
      selected.add(glyphId);
      queue.push(glyphId);
    }
  }

  for (let glyphId = queue.pop(); glyphId !== undefined; glyphId = queue.pop()) {
    const range = glyphRange(source, glyphId);
    if (!range) {
      continue;
    }
    if (
      range.end - range.start < GLYPH_HEADER_SIZE ||
      !fitsInView(source.view, range.start, GLYPH_HEADER_SIZE)
    ) {
      return fail(`glyph ${glyphId} is shorter than its ${GLYPH_HEADER_SIZE}-byte header`);
    }
    if (source.view.getInt16(range.start) > COMPOSITE_CONTOUR_COUNT) {
      continue;
    }

    const components = readCompositeComponents(source.view, range.start, range.end);
    if (!components) {
      return fail(`composite glyph ${glyphId} has a truncated component record`);
    }
    for (const component of components.glyphIds) {
      if (component >= source.numGlyphs) {
        return fail(
          `composite glyph ${glyphId} references glyph ${component}, which does not exist`,
        );
      }
      if (!selected.has(component)) {
        selected.add(component);
        queue.push(component);
      }
    }
  }

  return Result.ok([...selected].sort((left, right) => left - right));
};

type GlyfBuildResult = {
  readonly glyf: Uint8Array;
  /** `newGlyphIds.length + 1` offsets into `glyf`. */
  readonly locaOffsets: readonly number[];
};

/**
 * Copies the selected glyphs into a new `glyf`, renumbering the component
 * glyph index of every composite record on the way. That uint16 sits directly
 * after each component's flags word, and leaving it pointing at an old id is
 * the classic way a subset renders as garbage.
 */
const buildGlyf = (
  source: GlyphSource,
  glyphIds: readonly number[],
  glyphIdMap: ReadonlyMap<number, number>,
): Result<GlyfBuildResult, SubsetError> => {
  const parts: Uint8Array[] = [];
  const locaOffsets: number[] = [];
  let total = 0;

  for (const glyphId of glyphIds) {
    locaOffsets.push(total);
    const range = glyphRange(source, glyphId);
    if (!range) {
      continue;
    }

    const length = range.end - range.start;
    const data = new Uint8Array(alignUp(length));
    data.set(new Uint8Array(source.view.buffer, source.view.byteOffset + range.start, length));

    if (source.view.getInt16(range.start) <= COMPOSITE_CONTOUR_COUNT) {
      const components = readCompositeComponents(source.view, range.start, range.end);
      if (!components) {
        return fail(`composite glyph ${glyphId} has a truncated component record`);
      }

      const copied = new DataView(data.buffer);
      for (const offset of components.glyphIndexOffsets) {
        const oldComponent = source.view.getUint16(offset);
        const newComponent = glyphIdMap.get(oldComponent);
        if (newComponent === undefined) {
          panic(`component glyph ${oldComponent} escaped the subset closure`);
        }
        copied.setUint16(offset - range.start, newComponent);
      }
    }

    parts.push(data);
    total += data.byteLength;
  }
  locaOffsets.push(total);

  const glyf = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    glyf.set(part, at);
    at += part.byteLength;
  }
  return Result.ok({ glyf, locaOffsets });
};

const buildLoca = (offsets: readonly number[], longFormat: boolean): Uint8Array => {
  const entrySize = longFormat ? 4 : 2;
  const loca = new Uint8Array(offsets.length * entrySize);
  const view = new DataView(loca.buffer);
  for (const [index, offset] of offsets.entries()) {
    if (longFormat) {
      view.setUint32(index * entrySize, offset);
    } else {
      view.setUint16(index * entrySize, offset / 2);
    }
  }
  return loca;
};

type HmtxSourceOptions = {
  readonly view: DataView;
  readonly hmtx: SfntTableRecord;
  readonly numberOfHMetrics: number;
};

/**
 * Left side bearing of one glyph in the source font: from its longHorMetric
 * when it has one, else from the trailing int16 array. Returns 0 when the
 * table is too short, which is what a monospaced tail with no lsb array means.
 */
const sourceLeftSideBearing = (
  { view, hmtx, numberOfHMetrics }: HmtxSourceOptions,
  glyphId: number,
): number => {
  const at =
    glyphId < numberOfHMetrics
      ? hmtx.offset + glyphId * HMTX_METRIC_SIZE + 2
      : hmtx.offset + numberOfHMetrics * HMTX_METRIC_SIZE + (glyphId - numberOfHMetrics) * 2;
  if (at + 2 > hmtx.offset + hmtx.length || !fitsInView(view, at, 2)) {
    return 0;
  }
  return view.getInt16(at);
};

/** Sums a table as big-endian uint32s, zero-padded, modulo 2^32. */
const checksumOf = (bytes: Uint8Array): number => {
  let sum = 0;
  for (let at = 0; at < bytes.byteLength; at += 4) {
    const word =
      ((bytes[at] ?? 0) << 24) |
      ((bytes[at + 1] ?? 0) << 16) |
      ((bytes[at + 2] ?? 0) << 8) |
      (bytes[at + 3] ?? 0);
    sum = (sum + (word >>> 0)) % UINT32_MODULUS;
  }
  return sum >>> 0;
};

const searchParametersFor = (numTables: number) => {
  let entrySelector = 0;
  while (1 << (entrySelector + 1) <= numTables) {
    entrySelector += 1;
  }
  const searchRange = (1 << entrySelector) * SFNT_DIRECTORY_ENTRY_SIZE;
  return {
    searchRange,
    entrySelector,
    rangeShift: numTables * SFNT_DIRECTORY_ENTRY_SIZE - searchRange,
  };
};

type OutputTable = { readonly tag: string; readonly data: Uint8Array };

/**
 * Writes the sfnt and stamps `head.checkSumAdjustment`: the field is zero
 * while the file checksum is taken, then set to the magic constant minus that
 * sum, so a validator summing the finished file lands on the magic.
 */
const assembleSfnt = (tables: readonly OutputTable[]): Uint8Array => {
  const sorted = [...tables].sort((left, right) => (left.tag < right.tag ? -1 : 1));
  const directorySize = sorted.length * SFNT_DIRECTORY_ENTRY_SIZE;
  const bodySize = sorted.reduce((total, table) => total + alignUp(table.data.byteLength), 0);
  const output = new Uint8Array(SFNT_HEADER_SIZE + directorySize + bodySize);
  const view = new DataView(output.buffer);

  const { searchRange, entrySelector, rangeShift } = searchParametersFor(sorted.length);
  view.setUint32(0, SFNT_VERSION.trueType);
  view.setUint16(4, sorted.length);
  view.setUint16(6, searchRange);
  view.setUint16(8, entrySelector);
  view.setUint16(10, rangeShift);

  let entry = SFNT_HEADER_SIZE;
  let dataOffset = SFNT_HEADER_SIZE + directorySize;
  let headOffset = -1;
  for (const table of sorted) {
    for (let index = 0; index < table.tag.length; index++) {
      view.setUint8(entry + index, table.tag.charCodeAt(index));
    }
    view.setUint32(entry + 4, checksumOf(table.data));
    view.setUint32(entry + 8, dataOffset);
    view.setUint32(entry + 12, table.data.byteLength);
    output.set(table.data, dataOffset);
    if (table.tag === SFNT_TABLE.head) {
      headOffset = dataOffset;
    }
    dataOffset += alignUp(table.data.byteLength);
    entry += SFNT_DIRECTORY_ENTRY_SIZE;
  }

  if (headOffset >= 0) {
    view.setUint32(
      headOffset + HEAD.checkSumAdjustment,
      (CHECKSUM_ADJUSTMENT_MAGIC - checksumOf(output)) >>> 0,
    );
  }
  return output;
};

/** A copy of a source table, so edits never touch the caller's buffer. */
const copyTable = (directory: SfntDirectory, record: SfntTableRecord): Uint8Array =>
  new Uint8Array(
    directory.view.buffer.slice(
      directory.view.byteOffset + record.offset,
      directory.view.byteOffset + record.offset + record.length,
    ),
  );

const requireTable = (
  directory: SfntDirectory,
  tag: string,
  size: number,
): Result<SfntTableRecord, SubsetError> => {
  const record = directory.tables.get(tag);
  if (!record) {
    return fail(`font has no '${tag}' table, which a TrueType subset needs`);
  }
  if (record.length < size || !fitsInView(directory.view, record.offset, size)) {
    return fail(`'${tag}' table is ${record.length} bytes, needs at least ${size}`);
  }
  return Result.ok(record);
};

/**
 * Build a `glyf`-flavoured subset containing `glyphIds` plus the components
 * they reference transitively, plus .notdef.
 */
export const subsetTrueType = (
  font: SfntFont,
  glyphIds: ReadonlySet<number>,
): Result<TrueTypeSubset, SubsetError> => {
  if (font.isCff) {
    return fail(
      "cannot subset a CFF font: its outlines are charstrings, not glyf data; embed the font whole instead",
    );
  }

  const directoryResult = readSfntDirectory(font.bytes).mapError(
    (error) => new SubsetError({ message: error.message }),
  );
  if (directoryResult.isErr()) {
    return Result.err(directoryResult.error);
  }
  const directory = directoryResult.value;
  const { view } = directory;

  const headResult = requireTable(directory, SFNT_TABLE.head, HEAD.size);
  if (headResult.isErr()) {
    return Result.err(headResult.error);
  }
  const hheaResult = requireTable(directory, SFNT_TABLE.hhea, HHEA.size);
  if (hheaResult.isErr()) {
    return Result.err(hheaResult.error);
  }
  const maxpResult = requireTable(directory, SFNT_TABLE.maxp, MAXP.size);
  if (maxpResult.isErr()) {
    return Result.err(maxpResult.error);
  }
  const hmtxResult = requireTable(directory, SFNT_TABLE.hmtx, HMTX_METRIC_SIZE);
  if (hmtxResult.isErr()) {
    return Result.err(hmtxResult.error);
  }
  const glyfResult = requireTable(directory, SFNT_TABLE.glyf, 0);
  if (glyfResult.isErr()) {
    return Result.err(glyfResult.error);
  }
  const locaResult = requireTable(directory, SFNT_TABLE.loca, 0);
  if (locaResult.isErr()) {
    return Result.err(locaResult.error);
  }

  const head = headResult.value;
  const offsets = readLocaOffsets({
    view,
    loca: locaResult.value,
    numGlyphs: font.numGlyphs,
    longFormat: view.getInt16(head.offset + HEAD.indexToLocFormat) === 1,
  });
  if (!offsets) {
    return fail("'loca' is truncated or its offsets are not monotonic");
  }

  const source = { view, glyf: glyfResult.value, offsets, numGlyphs: font.numGlyphs };
  const selectedResult = closeOverComponents(source, glyphIds);
  if (selectedResult.isErr()) {
    return Result.err(selectedResult.error);
  }
  const selected = selectedResult.value;

  const glyphIdMap = new Map<number, number>();
  for (const [newId, oldId] of selected.entries()) {
    glyphIdMap.set(oldId, newId);
  }

  const glyfResultBuilt = buildGlyf(source, selected, glyphIdMap);
  if (glyfResultBuilt.isErr()) {
    return Result.err(glyfResultBuilt.error);
  }
  const { glyf, locaOffsets } = glyfResultBuilt.value;

  const longLoca = glyf.byteLength > MAX_SHORT_LOCA_GLYF_LENGTH;
  const loca = buildLoca(locaOffsets, longLoca);

  const newHead = copyTable(directory, head);
  const headView = new DataView(newHead.buffer);
  headView.setInt16(HEAD.indexToLocFormat, longLoca ? 1 : 0);
  headView.setUint32(HEAD.checkSumAdjustment, 0);

  const newHhea = copyTable(directory, hheaResult.value);
  new DataView(newHhea.buffer).setUint16(HHEA.numberOfHMetrics, selected.length);

  const newMaxp = copyTable(directory, maxpResult.value);
  new DataView(newMaxp.buffer).setUint16(MAXP.numGlyphs, selected.length);

  const sourceMetrics = {
    view,
    hmtx: hmtxResult.value,
    numberOfHMetrics: view.getUint16(hheaResult.value.offset + HHEA.numberOfHMetrics),
  };
  const newHmtx = new Uint8Array(selected.length * HMTX_METRIC_SIZE);
  const hmtxView = new DataView(newHmtx.buffer);
  for (const [newId, oldId] of selected.entries()) {
    hmtxView.setUint16(newId * HMTX_METRIC_SIZE, font.advanceWidthFor(oldId));
    hmtxView.setInt16(newId * HMTX_METRIC_SIZE + 2, sourceLeftSideBearing(sourceMetrics, oldId));
  }

  const tables: OutputTable[] = [
    { tag: SFNT_TABLE.head, data: newHead },
    { tag: SFNT_TABLE.hhea, data: newHhea },
    { tag: SFNT_TABLE.maxp, data: newMaxp },
    { tag: SFNT_TABLE.hmtx, data: newHmtx },
    { tag: SFNT_TABLE.loca, data: loca },
    { tag: SFNT_TABLE.glyf, data: glyf },
  ];
  for (const tag of VERBATIM_TABLES) {
    const record = directory.tables.get(tag);
    if (record) {
      tables.push({ tag, data: copyTable(directory, record) });
    }
  }

  return Result.ok({ bytes: assembleSfnt(tables), glyphIdMap });
};
