/**
 * Numbering instances the editor defines itself.
 *
 * A list command in a document with no list of the requested kind has nothing
 * to reference: the package may have no numbering part at all. The command
 * then defines one, an `w:abstractNum` built from {@link templateLevel}
 * plus the `w:num` naming it, as a word processor does when a list is
 * started in a document without lists.
 *
 * The definition lives in two places without being stored twice. While the
 * document is edited, the paragraphs that reference the instance carry the
 * rendering it resolved to (`listAbstractNumId`, `listNumFmt`,
 * `listMarkerTemplate`, `listStartOverride`, …), and those attrs travel with
 * the paragraph through undo, collaboration and every save path.
 * {@link completeListNumbering} turns them back into the definition wherever
 * the package's own numbering has no entry for the id. The rendering states
 * the customized format and start of each level a command minted.
 */

import { NUMBER_FORMATS } from "@stll/docx-core/model";

import type {
  AbstractNumbering,
  CounterFormat,
  LevelOverride,
  ListLevel,
  ListRendering,
  NumberFormat,
  NumberingDefinitions,
  NumberingInstance,
  ParagraphAlignment,
  ParagraphFormatting,
} from "../types/document";
import {
  createNumberingIdAllocator,
  mintNumberingId,
  type NumberingIdsInUse,
  numberingNumIdsTaken,
} from "./numberingIds";

/** The two kinds of list a command can create. */
export const LIST_KINDS = ["bullet", "numbered"] as const;
export type ListKind = (typeof LIST_KINDS)[number];

/** What a list level counts in and how its marker reads. */
export type ListLevelFormat = {
  numFmt: NumberFormat;
  /** `w:lvlText`, e.g. `%1.` or `%1)`. */
  lvlText: string;
};

/** The appearance and start of one level in a new list. */
export type NewListDefinition = {
  kind: ListKind;
  format?: ListLevelFormatPatch | undefined;
  start?: number | undefined;
  level?: number | undefined;
};

const INDENT_STEP_TWIPS = 720;
const HANGING_TWIPS = 360;
const ROMAN_HANGING_TWIPS = 180;
const LEVEL_COUNT = 9;

const levelIndent = (ilvl: number, hanging: number): ParagraphFormatting => ({
  indentLeft: INDENT_STEP_TWIPS * (ilvl + 1),
  indentFirstLine: -hanging,
  hangingIndent: true,
});

type NumberedCycleEntry = {
  numFmt: NumberFormat;
  lvlJc: ParagraphAlignment;
  hanging: number;
};

/** The customary numbered-list cycle: `1.`, `a.`, `i.`, then again one level deeper. */
const NUMBERED_CYCLE: readonly [
  NumberedCycleEntry,
  NumberedCycleEntry,
  NumberedCycleEntry,
] = [
  { numFmt: "decimal", lvlJc: "left", hanging: HANGING_TWIPS },
  { numFmt: "lowerLetter", lvlJc: "left", hanging: HANGING_TWIPS },
  { numFmt: "lowerRoman", lvlJc: "right", hanging: ROMAN_HANGING_TWIPS },
];

type BulletCycleEntry = { text: string; font: string | null };

/** The customary bullet cycle: a round bullet, an open circle, a square. */
const BULLET_CYCLE: readonly [
  BulletCycleEntry,
  BulletCycleEntry,
  BulletCycleEntry,
] = [
  { text: "\u2022", font: null },
  { text: "o", font: "Courier New" },
  { text: "\u25AA", font: null },
];

const cycleEntry = <T>(cycle: readonly [T, T, T], ilvl: number): T =>
  cycle.at(ilvl % cycle.length) ?? cycle[0];

const numberedLevel = (ilvl: number): ListLevel => {
  const { numFmt, lvlJc, hanging } = cycleEntry(NUMBERED_CYCLE, ilvl);
  return {
    ilvl,
    start: 1,
    numFmt,
    lvlText: `%${ilvl + 1}.`,
    lvlJc,
    pPr: levelIndent(ilvl, hanging),
  };
};

const bulletLevel = (ilvl: number): ListLevel => {
  const { text, font } = cycleEntry(BULLET_CYCLE, ilvl);
  return {
    ilvl,
    start: 1,
    numFmt: "bullet",
    lvlText: text,
    lvlJc: "left",
    pPr: levelIndent(ilvl, HANGING_TWIPS),
    ...(font === null
      ? {}
      : { rPr: { fontFamily: { ascii: font, hAnsi: font } } }),
  };
};

const templateLevel = (kind: ListKind, ilvl: number): ListLevel =>
  kind === "bullet" ? bulletLevel(ilvl) : numberedLevel(ilvl);

const numberFormatKind = (numFmt: NumberFormat): ListKind =>
  numFmt === "bullet" ? "bullet" : "numbered";

const withLevelFormat = (
  template: ListLevel,
  format: ListLevelFormatPatch,
): ListLevel => ({
  ...template,
  ...(format.numFmt === undefined ? {} : { numFmt: format.numFmt }),
  ...(format.lvlText === undefined ? {} : { lvlText: format.lvlText }),
});

type ListLevelFormatPatch = {
  numFmt?: NumberFormat | undefined;
  lvlText?: string | undefined;
};

type MintAbstractNumberingOptions = {
  abstractNumId: number;
  kind: ListKind;
  format?: ListLevelFormatPatch | undefined;
  level?: number | undefined;
};

const mintAbstractNumbering = ({
  abstractNumId,
  kind,
  format,
  level = 0,
}: MintAbstractNumberingOptions): AbstractNumbering => ({
  abstractNumId,
  multiLevelType: "hybridMultilevel",
  levels: Array.from({ length: LEVEL_COUNT }, (_, ilvl) => {
    if (format === undefined || ilvl !== level)
      return templateLevel(kind, ilvl);
    const levelKind =
      format.numFmt === undefined ? kind : numberFormatKind(format.numFmt);
    return withLevelFormat(templateLevel(levelKind, ilvl), format);
  }),
});

const EMPTY_DEFINITIONS: NumberingDefinitions = { abstractNums: [], nums: [] };

/** A fresh `w:numId` for `definitions`; zero is reserved for "no numbering". */
const nextNumId = (
  definitions: NumberingDefinitions | null | undefined,
  elsewhere: NumberingIdsInUse,
): number =>
  mintNumberingId({
    kind: "num",
    existingIds: [...numberingNumIdsTaken(definitions), ...elsewhere.numIds],
  });

/** A fresh `w:abstractNumId` for `definitions`. */
const nextAbstractNumId = (
  definitions: NumberingDefinitions | null | undefined,
  elsewhere: NumberingIdsInUse,
): number =>
  mintNumberingId({
    kind: "abstract",
    existingIds: [
      ...(definitions?.abstractNums ?? []).map(
        ({ abstractNumId }) => abstractNumId,
      ),
      ...elsewhere.abstractNumIds,
    ],
  });

type MintedListInstance = {
  definitions: NumberingDefinitions;
  numId: number;
};

/**
 * Define a new list: a new `w:abstractNum` of the requested kind and the
 * `w:num` naming it. `definitions` must already hold every instance the
 * story references (see {@link completeListNumbering}), and `elsewhere` the
 * ids the package's other stories define, so the new ids collide with none of
 * them.
 */
export const mintListInstance = (
  definitions: NumberingDefinitions | null | undefined,
  { kind, format, start, level = 0 }: NewListDefinition,
  elsewhere: NumberingIdsInUse,
): MintedListInstance => {
  const current = definitions ?? EMPTY_DEFINITIONS;
  const numId = nextNumId(current, elsewhere);
  const abstractNumId = nextAbstractNumId(current, elsewhere);
  const levelOverrides =
    start === undefined || start === 1
      ? undefined
      : [{ ilvl: level, startOverride: start }];
  const instance: NumberingInstance = {
    numId,
    abstractNumId,
    ...(levelOverrides ? { levelOverrides } : {}),
  };
  return {
    numId,
    definitions: {
      ...current,
      abstractNums: [
        ...current.abstractNums,
        mintAbstractNumbering({ abstractNumId, kind, format, level }),
      ],
      nums: [...current.nums, instance],
    },
  };
};

type RestartListInstanceOptions = {
  abstractNumId: number;
  ilvl: number;
  start: number;
};

/**
 * A new `w:num` over an existing `w:abstractNum`, with a `w:startOverride` at
 * `ilvl`: *Restart Numbering* and *Set Numbering Value*. The override is what
 * makes a consumer count the instance on its own rather than continue the
 * other instances of the same definition.
 */
export const restartListInstance = (
  definitions: NumberingDefinitions | null | undefined,
  { abstractNumId, ilvl, start }: RestartListInstanceOptions,
  elsewhere: NumberingIdsInUse,
): MintedListInstance => {
  const current = definitions ?? EMPTY_DEFINITIONS;
  const numId = nextNumId(current, elsewhere);
  const instance: NumberingInstance = {
    numId,
    abstractNumId,
    levelOverrides: [{ ilvl, startOverride: start }],
  };
  return {
    numId,
    definitions: { ...current, nums: [...current.nums, instance] },
  };
};

/** One paragraph's numbering reference and the rendering it resolved to. */
export type ListInstanceReference = {
  numId: number;
  ilvl: number;
  rendering: ListRendering;
};

const NUMBER_FORMAT_SET: ReadonlySet<string> = new Set(NUMBER_FORMATS);

const isNumberFormat = (
  format: CounterFormat | undefined,
): format is NumberFormat =>
  format !== undefined && NUMBER_FORMAT_SET.has(format);

const rebuildAbstract = (
  abstractNumId: number,
  references: readonly ListInstanceReference[],
): AbstractNumbering => {
  const active = new Map(
    references.map(({ ilvl, rendering }) => [ilvl, rendering] as const),
  );
  const declared = references.find(({ rendering }) => rendering.levelNumFmts)
    ?.rendering.levelNumFmts;
  const fallbackKind = references.some(({ rendering }) => rendering.isBullet)
    ? "bullet"
    : "numbered";
  return {
    abstractNumId,
    multiLevelType: "hybridMultilevel",
    levels: Array.from({ length: LEVEL_COUNT }, (_, ilvl) => {
      const rendering = active.get(ilvl);
      const format = rendering?.numFmt ?? declared?.at(ilvl);
      const kind = isNumberFormat(format)
        ? numberFormatKind(format)
        : fallbackKind;
      return withLevelFormat(templateLevel(kind, ilvl), {
        numFmt: isNumberFormat(format) ? format : undefined,
        lvlText: rendering?.markerTemplate,
      });
    }),
  };
};

const startOverridesOf = (
  references: readonly ListInstanceReference[],
): LevelOverride[] => {
  const byLevel = new Map<number, number>();
  for (const { ilvl, rendering } of references) {
    if (rendering.startOverride !== undefined && !byLevel.has(ilvl)) {
      byLevel.set(ilvl, rendering.startOverride);
    }
  }
  return [...byLevel.entries()]
    .toSorted(([left], [right]) => left - right)
    .map(([ilvl, startOverride]) => ({ ilvl, startOverride }));
};

/**
 * Define every instance `references` name that `definitions` does not.
 *
 * A reference the editor resolved carries its rendering, and a rendering is
 * what a minted instance leaves behind; a reference without one (a dangling
 * id read from a file) has no rendering and so is never defined here. Returns
 * `definitions` itself when nothing is missing.
 */
export const completeListNumbering = (
  definitions: NumberingDefinitions | undefined,
  references: Iterable<ListInstanceReference>,
): NumberingDefinitions | undefined => {
  const definedNums = new Set(
    (definitions?.nums ?? []).map(({ numId }) => numId),
  );
  const missing = new Map<number, ListInstanceReference[]>();
  for (const reference of references) {
    if (definedNums.has(reference.numId)) {
      continue;
    }
    const group = missing.get(reference.numId);
    if (group) {
      group.push(reference);
    } else {
      missing.set(reference.numId, [reference]);
    }
  }
  if (missing.size === 0) {
    return definitions;
  }

  const current = definitions ?? EMPTY_DEFINITIONS;
  const abstractIds = new Set(
    current.abstractNums.map(({ abstractNumId }) => abstractNumId),
  );
  const statedAbstractIds = [...missing.values()].flatMap((group) =>
    group.flatMap(({ rendering }) =>
      rendering.abstractNumId === undefined ? [] : [rendering.abstractNumId],
    ),
  );
  const missingAbstractIds = createNumberingIdAllocator("abstract", [
    ...abstractIds,
    ...statedAbstractIds,
  ]);
  const abstractNums: AbstractNumbering[] = [];
  const nums: NumberingInstance[] = [];

  for (const [numId, group] of missing) {
    const abstractNumId =
      group.find(({ rendering }) => rendering.abstractNumId !== undefined)
        ?.rendering.abstractNumId ?? missingAbstractIds.next();
    if (!abstractIds.has(abstractNumId)) {
      abstractIds.add(abstractNumId);
      abstractNums.push(rebuildAbstract(abstractNumId, group));
    }
    const levelOverrides = startOverridesOf(group);
    nums.push({
      numId,
      abstractNumId,
      ...(levelOverrides.length > 0 ? { levelOverrides } : {}),
    });
  }

  return {
    ...current,
    abstractNums: [...current.abstractNums, ...abstractNums],
    nums: [...current.nums, ...nums],
  };
};
