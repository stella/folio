/**
 * Numbering instances the editor defines itself.
 *
 * A list command in a document with no list of the requested kind has nothing
 * to reference: the package may have no numbering part at all. The command
 * then defines one, an `w:abstractNum` built from {@link defaultListLevels}
 * plus the `w:num` naming it, as a word processor does when a list is
 * started in a document without lists.
 *
 * The definition lives in two places without being stored twice. While the
 * document is edited, the paragraphs that reference the instance carry the
 * rendering it resolved to (`listAbstractNumId`, `listNumFmt`,
 * `listMarkerTemplate`, `listStartOverride`, …), and those attrs travel with
 * the paragraph through undo, collaboration and every save path.
 * {@link completeListNumbering} turns them back into the definition wherever
 * the package's own numbering has no entry for the id. Because a minted
 * definition is the template plus level-0 facts that the rendering states, the
 * definition it rebuilds is the one the command minted.
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

/** The two kinds of list a command can create. */
export const LIST_KINDS = ["bullet", "numbered"] as const;
export type ListKind = (typeof LIST_KINDS)[number];

/** What a numbered list's first level counts in and how its marker reads. */
export type ListLevelFormat = {
  numFmt: NumberFormat;
  /** `w:lvlText` for level 0, e.g. `%1.` or `%1)`. */
  lvlText: string;
};

/** The first-level appearance and start a new list is minted with. */
type NewListDefinition = {
  kind: ListKind;
  /** Level-0 format of a numbered list; the template's `%1.` decimal otherwise. */
  format?: ListLevelFormat | undefined;
  /** The value the first item shows, stated as a level-0 `w:startOverride`. */
  start?: number | undefined;
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
const NUMBERED_CYCLE: readonly [NumberedCycleEntry, NumberedCycleEntry, NumberedCycleEntry] = [
  { numFmt: "decimal", lvlJc: "left", hanging: HANGING_TWIPS },
  { numFmt: "lowerLetter", lvlJc: "left", hanging: HANGING_TWIPS },
  { numFmt: "lowerRoman", lvlJc: "right", hanging: ROMAN_HANGING_TWIPS },
];

type BulletCycleEntry = { text: string; font: string | null };

/** The customary bullet cycle: a round bullet, an open circle, a square. */
const BULLET_CYCLE: readonly [BulletCycleEntry, BulletCycleEntry, BulletCycleEntry] = [
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
    ...(font === null ? {} : { rPr: { fontFamily: { ascii: font, hAnsi: font } } }),
  };
};

/** The nine levels a list of `kind` is minted with. */
const defaultListLevels = (kind: ListKind): ListLevel[] =>
  Array.from({ length: LEVEL_COUNT }, (_, ilvl) =>
    kind === "bullet" ? bulletLevel(ilvl) : numberedLevel(ilvl),
  );

type MintAbstractNumberingOptions = {
  abstractNumId: number;
  kind: ListKind;
  format?: ListLevelFormat | undefined;
};

const mintAbstractNumbering = ({
  abstractNumId,
  kind,
  format,
}: MintAbstractNumberingOptions): AbstractNumbering => {
  const levels = defaultListLevels(kind);
  const first = levels[0];
  if (first && kind === "numbered" && format) {
    levels[0] = { ...first, numFmt: format.numFmt, lvlText: format.lvlText };
  }
  return { abstractNumId, multiLevelType: "hybridMultilevel", levels };
};

/** The template counts from one, so only another first value needs an override. */
const levelZeroStart = (start: number | undefined): LevelOverride[] | undefined =>
  start === undefined || start === 1 ? undefined : [{ ilvl: 0, startOverride: start }];

const EMPTY_DEFINITIONS: NumberingDefinitions = { abstractNums: [], nums: [] };

const nextId = (ids: Iterable<number>, floor: number): number => {
  let next = floor;
  for (const id of ids) {
    next = Math.max(next, id + 1);
  }
  return next;
};

/** A fresh `w:numId` for `definitions`; zero is reserved for "no numbering". */
const nextNumId = (definitions: NumberingDefinitions | null | undefined): number =>
  nextId(
    (definitions?.nums ?? []).map(({ numId }) => numId),
    1,
  );

/** A fresh `w:abstractNumId` for `definitions`. */
const nextAbstractNumId = (definitions: NumberingDefinitions | null | undefined): number =>
  nextId(
    (definitions?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
    0,
  );

type MintedListInstance = {
  definitions: NumberingDefinitions;
  numId: number;
};

/**
 * Define a new list: a new `w:abstractNum` of the requested kind and the
 * `w:num` naming it. `definitions` must already hold every instance the
 * document references (see {@link completeListNumbering}), so the new ids
 * collide with none of them.
 */
export const mintListInstance = (
  definitions: NumberingDefinitions | null | undefined,
  { kind, format, start }: NewListDefinition,
): MintedListInstance => {
  const current = definitions ?? EMPTY_DEFINITIONS;
  const numId = nextNumId(current);
  const abstractNumId = nextAbstractNumId(current);
  const levelOverrides = levelZeroStart(start);
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
        mintAbstractNumbering({ abstractNumId, kind, format }),
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
): MintedListInstance => {
  const current = definitions ?? EMPTY_DEFINITIONS;
  const numId = nextNumId(current);
  const instance: NumberingInstance = {
    numId,
    abstractNumId,
    levelOverrides: [{ ilvl, startOverride: start }],
  };
  return { numId, definitions: { ...current, nums: [...current.nums, instance] } };
};

/** One paragraph's numbering reference and the rendering it resolved to. */
export type ListInstanceReference = {
  numId: number;
  ilvl: number;
  rendering: ListRendering;
};

const NUMBER_FORMAT_SET: ReadonlySet<string> = new Set(NUMBER_FORMATS);

const isNumberFormat = (format: CounterFormat | undefined): format is NumberFormat =>
  format !== undefined && NUMBER_FORMAT_SET.has(format);

const levelZeroFormat = (
  references: readonly ListInstanceReference[],
): ListLevelFormat | undefined => {
  const first = references.find(({ ilvl }) => ilvl === 0)?.rendering;
  const numFmt =
    first?.numFmt ??
    references.find(({ rendering }) => rendering.levelNumFmts?.[0])?.rendering.levelNumFmts?.[0];
  if (!isNumberFormat(numFmt)) {
    return undefined;
  }
  return { numFmt, lvlText: first?.markerTemplate ?? "%1." };
};

const startOverridesOf = (references: readonly ListInstanceReference[]): LevelOverride[] => {
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
  const definedNums = new Set((definitions?.nums ?? []).map(({ numId }) => numId));
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
  const abstractIds = new Set(current.abstractNums.map(({ abstractNumId }) => abstractNumId));
  const statedAbstractIds = [...missing.values()].flatMap((group) =>
    group.flatMap(({ rendering }) =>
      rendering.abstractNumId === undefined ? [] : [rendering.abstractNumId],
    ),
  );
  let freeAbstractId = nextId([...abstractIds, ...statedAbstractIds], 0);
  const abstractNums: AbstractNumbering[] = [];
  const nums: NumberingInstance[] = [];

  for (const [numId, group] of missing) {
    const abstractNumId =
      group.find(({ rendering }) => rendering.abstractNumId !== undefined)?.rendering
        .abstractNumId ?? freeAbstractId++;
    if (!abstractIds.has(abstractNumId)) {
      abstractIds.add(abstractNumId);
      const kind: ListKind = group.some(({ rendering }) => rendering.isBullet)
        ? "bullet"
        : "numbered";
      abstractNums.push(
        mintAbstractNumbering({ abstractNumId, kind, format: levelZeroFormat(group) }),
      );
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
