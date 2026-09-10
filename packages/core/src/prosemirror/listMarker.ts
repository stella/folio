import { formatOoxmlCounter } from "../docx/ooxmlCounterFormatter";
import { convertBulletToUnicode } from "../docx/bulletMarkers";
import type { NumberFormat } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

export type ListCounterState = {
  counters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  seenLevels: Set<string>;
  restartedNumIds?: Set<number>;
  siblingNumIdsByAbstractNumId?: Map<number, Set<number>>;
  previousList?: {
    abstractNumId: number | null;
    fromStyle: boolean;
    numId: number | null;
  };
  lastAdvancedNumId?: number;
};

export type ListCounterStreams = {
  final: ListCounterState;
  original: ListCounterState;
};

export type ListCounterStream = keyof ListCounterStreams;

export type AdvancedListMarker = {
  marker: string | null;
  counterAttrs: ParagraphAttrs;
  counterState: ListCounterState;
  stream: ListCounterStream;
};

export type VisibleListMarker = {
  marker: string | null;
  counterAttrs: ParagraphAttrs;
  counterState: ListCounterState;
  stream: ListCounterStream;
  advances: readonly AdvancedListMarker[];
};

export type ResolvedListComponent = {
  level: number;
  relativeStart: number;
  start: number;
  end: number;
};

export type ResolvedListTemplate = {
  value: string;
  components: readonly ResolvedListComponent[];
};

export const MAX_LIST_LEVEL = 8;

export const LIST_RENDERING_ATTR_KEYS = [
  "listIsBullet",
  "listIsLegal",
  "listNumFmt",
  "listMarker",
  "listMarkerTemplate",
  "listMarkerHidden",
  "listMarkerFormatting",
  "listMarkerAlignment",
  "listMarkerSuffix",
  "listMarkerAllCaps",
  "listImplicitChildLevelAdvances",
  "listMarkerSecondSlotOffsetTwips",
  "listLevelNumFmts",
  "listLevelStarts",
  "listAbstractNumId",
  "listStartOverride",
] as const satisfies readonly (keyof ParagraphAttrs)[];

export const CLEARED_LIST_RENDERING_ATTRS = Object.freeze({
  listIsBullet: null,
  listIsLegal: null,
  listNumFmt: null,
  listMarker: null,
  listMarkerTemplate: null,
  listMarkerHidden: null,
  listMarkerFormatting: null,
  listMarkerAlignment: null,
  listMarkerSuffix: null,
  listMarkerAllCaps: null,
  listImplicitChildLevelAdvances: null,
  listMarkerSecondSlotOffsetTwips: null,
  listLevelNumFmts: null,
  listLevelStarts: null,
  listAbstractNumId: null,
  listStartOverride: null,
});

export function createListCounterState(): ListCounterState {
  return {
    counters: new Map(),
    abstractCounters: new Map(),
    seenLevels: new Set(),
    restartedNumIds: new Set(),
    siblingNumIdsByAbstractNumId: new Map(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
  };
}

export function cloneListCounterState(state: ListCounterState): ListCounterState {
  const counters = new Map<number, number[]>();
  for (const [numId, values] of state.counters) {
    counters.set(numId, [...values]);
  }
  const abstractCounters = new Map<number, number[]>();
  for (const [abstractNumId, values] of state.abstractCounters) {
    abstractCounters.set(abstractNumId, [...values]);
  }
  return {
    counters,
    abstractCounters,
    seenLevels: new Set(state.seenLevels),
    restartedNumIds: new Set(state.restartedNumIds ?? []),
    siblingNumIdsByAbstractNumId: new Map(
      [...(state.siblingNumIdsByAbstractNumId ?? [])].map(([abstractNumId, numIds]) => [
        abstractNumId,
        new Set(numIds),
      ]),
    ),
    previousList: {
      ...(state.previousList ?? { abstractNumId: null, fromStyle: false, numId: null }),
    },
    ...(state.lastAdvancedNumId !== undefined
      ? { lastAdvancedNumId: state.lastAdvancedNumId }
      : {}),
  };
}

const isListNumPr = (
  value: ParagraphAttrs["numPr"] | null | undefined,
): value is NonNullable<ParagraphAttrs["numPr"]> => value !== undefined && value !== null;

const sameListNumPr = (
  left: NonNullable<ParagraphAttrs["numPr"]>,
  right: NonNullable<ParagraphAttrs["numPr"]>,
): boolean => left.numId === right.numId && left.ilvl === right.ilvl;

function previousListAttrs(attrs: ParagraphAttrs): ParagraphAttrs | null {
  const change = attrs._propertyChanges?.find(({ previousFormatting }) => {
    const previousNumPr = previousFormatting?.numPr;
    return (
      isListNumPr(previousNumPr) && (!attrs.numPr || !sameListNumPr(previousNumPr, attrs.numPr))
    );
  });
  const previous = change?.previousFormatting;
  if (!previous || !isListNumPr(previous.numPr)) {
    return null;
  }
  return {
    numPr: previous.numPr,
    ...(previous.listIsBullet !== undefined ? { listIsBullet: previous.listIsBullet } : {}),
    ...(previous.listIsLegal !== undefined ? { listIsLegal: previous.listIsLegal } : {}),
    ...(previous.listNumFmt !== undefined ? { listNumFmt: previous.listNumFmt } : {}),
    ...(previous.listMarker !== undefined ? { listMarker: previous.listMarker } : {}),
    ...(previous.listMarkerTemplate !== undefined
      ? { listMarkerTemplate: previous.listMarkerTemplate }
      : {}),
    ...(previous.listMarkerHidden !== undefined
      ? { listMarkerHidden: previous.listMarkerHidden }
      : {}),
    ...(previous.listMarkerFormatting !== undefined
      ? { listMarkerFormatting: previous.listMarkerFormatting }
      : {}),
    ...(previous.listMarkerAlignment !== undefined
      ? { listMarkerAlignment: previous.listMarkerAlignment }
      : {}),
    ...(previous.listMarkerSuffix !== undefined
      ? { listMarkerSuffix: previous.listMarkerSuffix }
      : {}),
    ...(previous.listMarkerAllCaps !== undefined
      ? { listMarkerAllCaps: previous.listMarkerAllCaps }
      : {}),
    ...(previous.listImplicitChildLevelAdvances !== undefined
      ? { listImplicitChildLevelAdvances: previous.listImplicitChildLevelAdvances }
      : {}),
    ...(previous.listMarkerSecondSlotOffsetTwips !== undefined
      ? { listMarkerSecondSlotOffsetTwips: previous.listMarkerSecondSlotOffsetTwips }
      : {}),
    ...(previous.listLevelNumFmts !== undefined
      ? { listLevelNumFmts: previous.listLevelNumFmts }
      : {}),
    ...(previous.listLevelStarts !== undefined
      ? { listLevelStarts: previous.listLevelStarts }
      : {}),
    ...(previous.listAbstractNumId !== undefined
      ? { listAbstractNumId: previous.listAbstractNumId }
      : {}),
    ...(previous.listStartOverride !== undefined
      ? { listStartOverride: previous.listStartOverride }
      : {}),
  };
}

/**
 * Advance the counter stream that paints this paragraph's marker. Normal
 * paragraphs advance both final and original streams; inserted/deleted list
 * paragraphs advance only the document view in which they exist.
 */
export function advanceVisibleListMarker(
  attrs: ParagraphAttrs,
  streams: ListCounterStreams,
): VisibleListMarker {
  const advances: AdvancedListMarker[] = [];
  const advance = (counterAttrs: ParagraphAttrs, stream: ListCounterStream): AdvancedListMarker => {
    const counterState = streams[stream];
    const advanced = {
      marker: advanceListMarker(counterAttrs, counterState),
      counterAttrs,
      counterState,
      stream,
    };
    advances.push(advanced);
    return advanced;
  };
  const previous = previousListAttrs(attrs);
  if (!attrs.numPr) {
    const counterAttrs = previous ?? attrs;
    const visible = advance(counterAttrs, previous ? "original" : "final");
    return { ...visible, advances };
  }

  if (attrs.pPrMark?.kind === "del") {
    const visible = advance(attrs, "original");
    return { ...visible, advances };
  }

  const numberingWasAdded = attrs._propertyChanges?.some(
    ({ previousFormatting }) =>
      previousFormatting &&
      Object.hasOwn(previousFormatting, "numPr") &&
      previousFormatting.numPr == null,
  );
  const numberingChanged =
    previous?.numPr !== undefined && !sameListNumPr(previous.numPr, attrs.numPr);
  const visible = advance(attrs, "final");
  if (attrs.pPrMark?.kind !== "ins" && !numberingWasAdded && !numberingChanged) {
    advance(attrs, "original");
  } else if (numberingChanged && previous) {
    advance(previous, "original");
  }
  return { ...visible, advances };
}

export function formatCounter(value: number, format: NumberFormat | undefined): string {
  return formatOoxmlCounter(value, format);
}

type ResolveListTemplateOptions = {
  template: string;
  counters: number[];
  levelFormats?: NumberFormat[] | undefined;
  forceDecimal?: boolean | undefined;
};

export function resolveListTemplate(options: ResolveListTemplateOptions): string {
  return resolveListTemplateWithComponents(options).value;
}

export function resolveListTemplateWithComponents({
  template,
  counters,
  levelFormats,
  forceDecimal = false,
}: ResolveListTemplateOptions): ResolvedListTemplate {
  const components: ResolvedListComponent[] = [];
  let value = "";
  let sourceEnd = 0;
  for (const match of template.matchAll(/%(?<digit>\d)(?<punct>[.):\]])?/gu)) {
    const sourceStart = match.index;
    const relativeStart = value.length;
    value += template.slice(sourceEnd, sourceStart);
    sourceEnd = sourceStart + match[0].length;

    const index = Number.parseInt(match.groups?.["digit"] ?? "", 10) - 1;
    if (index < 0) {
      continue;
    }
    const counter = counters[index];
    if (counter === undefined || Number.isNaN(counter)) {
      continue;
    }
    const formatted = formatCounter(counter, forceDecimal ? "decimal" : levelFormats?.[index]);
    if (!formatted) {
      continue;
    }
    const start = value.length;
    value += formatted;
    components.push({ level: index, relativeStart, start, end: value.length });
    value += match.groups?.["punct"] ?? "";
  }
  value += template.slice(sourceEnd);
  return { value, components };
}

function getLastListCounters(state: ListCounterState): number[] | undefined {
  return state.lastAdvancedNumId === undefined
    ? undefined
    : state.counters.get(state.lastAdvancedNumId);
}

function formatNumberedMarker(counters: number[], level: number): string {
  const parts: number[] = [];
  for (let index = 0; index <= level; index += 1) {
    const value = counters[index] ?? 0;
    if (!Number.isFinite(value) || value <= 0) {
      break;
    }
    parts.push(value);
  }
  return parts.length === 0 ? "1." : `${parts.join(".")}.`;
}

export function advanceListMarker(attrs: ParagraphAttrs, state: ListCounterState): string | null {
  const markerTemplate = attrs.listMarkerTemplate ?? attrs.listMarker;
  const level = attrs.numPr?.ilvl ?? 0;
  if (!Number.isInteger(level) || level < 0 || level > MAX_LIST_LEVEL) {
    return null;
  }
  const numId = attrs.numPr?.numId;
  if (numId === undefined || numId === 0) {
    let marker: string | null = null;
    if (markerTemplate?.includes("%") && !attrs.listIsBullet) {
      const counters = getLastListCounters(state);
      if (counters) {
        marker = resolveListTemplate({
          template: markerTemplate,
          counters,
          levelFormats: attrs.listLevelNumFmts,
          forceDecimal: attrs.listIsLegal,
        });
      }
    }
    state.previousList = { abstractNumId: null, fromStyle: false, numId: null };
    return marker;
  }

  if (attrs.listIsBullet) {
    state.previousList = { abstractNumId: null, fromStyle: false, numId: null };
    return convertBulletToUnicode(attrs.listMarker ?? markerTemplate ?? "");
  }

  const firstInstanceEncounter = !state.counters.has(numId);
  const counters = state.counters.get(numId) ?? Array.from({ length: 9 }, () => Number.NaN);
  const abstractNumId = attrs.listAbstractNumId;
  const previousList = state.previousList ?? {
    abstractNumId: null,
    fromStyle: false,
    numId: null,
  };
  const restartedNumIds = (state.restartedNumIds ??= new Set());
  const siblingNumIdsByAbstractNumId = (state.siblingNumIdsByAbstractNumId ??= new Map());
  const latestAbstractCounters =
    abstractNumId === undefined ? undefined : state.abstractCounters.get(abstractNumId);
  if (firstInstanceEncounter && abstractNumId !== undefined) {
    const siblingNumIds = siblingNumIdsByAbstractNumId.get(abstractNumId) ?? new Set();
    siblingNumIds.add(numId);
    siblingNumIdsByAbstractNumId.set(abstractNumId, siblingNumIds);
  }
  const styleNumbering = attrs.numPrFromStyle;
  const resumesRestartedInstance =
    firstInstanceEncounter &&
    attrs.listStartOverride == null &&
    abstractNumId !== undefined &&
    previousList.abstractNumId === abstractNumId &&
    previousList.numId !== null &&
    previousList.numId !== numId &&
    restartedNumIds.has(previousList.numId);
  const resumesStyleInstance =
    firstInstanceEncounter &&
    !styleNumbering &&
    attrs.listStartOverride == null &&
    abstractNumId !== undefined &&
    previousList.abstractNumId === abstractNumId &&
    previousList.fromStyle;
  let resumedAbstractCounters: number[] | undefined;
  if (
    latestAbstractCounters &&
    (styleNumbering || resumesRestartedInstance || resumesStyleInstance)
  ) {
    for (let index = 0; index < counters.length; index += 1) {
      counters[index] = latestAbstractCounters[index] ?? Number.NaN;
    }
    resumedAbstractCounters = latestAbstractCounters;
  }
  if (
    attrs.listStartOverride != null ||
    resumesRestartedInstance ||
    (previousList.numId === numId && restartedNumIds.has(numId))
  ) {
    restartedNumIds.add(numId);
  }
  if (level > 0) {
    if (counters.slice(0, level).every((counter) => !Number.isFinite(counter))) {
      for (let index = 0; index < level; index += 1) {
        const latestCounter = latestAbstractCounters?.[index];
        counters[index] =
          latestCounter !== undefined && Number.isFinite(latestCounter)
            ? latestCounter
            : (attrs.listLevelStarts?.[index] ?? 1);
      }
    }
  }

  const seenKey = `${numId}:${level}`;
  const firstEncounter = !state.seenLevels.has(seenKey);
  if (firstEncounter) {
    state.seenLevels.add(seenKey);
    if (attrs.listStartOverride != null) {
      counters[level] = attrs.listStartOverride - 1;
    }
  }
  if (!Number.isFinite(counters[level])) {
    counters[level] = (attrs.listLevelStarts?.[level] ?? 1) - 1;
  }

  counters[level] = (counters[level] ?? 0) + 1;
  for (let index = level + 1; index < counters.length; index += 1) {
    counters[index] = Number.NaN;
  }
  const childAdvances = attrs.listImplicitChildLevelAdvances ?? 0;
  if (childAdvances > 0 && level + 1 < counters.length) {
    const childCounter = counters[level + 1];
    counters[level + 1] =
      (childCounter === undefined || !Number.isFinite(childCounter) ? 0 : childCounter) +
      childAdvances;
  }
  state.counters.set(numId, counters);
  state.lastAdvancedNumId = numId;
  if (abstractNumId !== undefined) {
    if (resumedAbstractCounters) {
      const siblingNumIds = siblingNumIdsByAbstractNumId.get(abstractNumId);
      for (const otherNumId of siblingNumIds ?? []) {
        const otherCounters = state.counters.get(otherNumId);
        if (
          otherNumId === numId ||
          !otherCounters ||
          !otherCounters.every((value, index) => Object.is(value, resumedAbstractCounters[index]))
        ) {
          continue;
        }
        for (let index = 0; index < otherCounters.length; index += 1) {
          otherCounters[index] = counters[index] ?? Number.NaN;
        }
      }
    }
    state.abstractCounters.set(abstractNumId, [...counters]);
  }
  state.previousList = {
    abstractNumId: abstractNumId ?? null,
    fromStyle: Boolean(styleNumbering),
    numId,
  };

  const levelFormats =
    attrs.listLevelNumFmts ?? (attrs.listNumFmt ? [attrs.listNumFmt] : undefined);
  if (markerTemplate?.includes("%")) {
    return resolveListTemplate({
      template: markerTemplate,
      counters,
      levelFormats,
      forceDecimal: attrs.listIsLegal,
    });
  }
  if (attrs.listMarker) {
    return attrs.listMarker;
  }
  const levelFormat = levelFormats?.[level] ?? attrs.listNumFmt;
  if (levelFormat === "none" || attrs.listMarker === "") {
    return null;
  }
  return formatNumberedMarker(counters, level);
}
