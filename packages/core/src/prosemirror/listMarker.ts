import { resolveListTemplate } from "../docx/listTemplate";

export { resolveListTemplate, resolveListTemplateWithComponents } from "../docx/listTemplate";
export type { ResolvedListComponent, ResolvedListTemplate } from "../docx/listTemplate";
import { formatOoxmlCounter } from "../docx/ooxmlCounterFormatter";
import { bulletMarkerFontName, convertBulletToUnicode } from "../docx/bulletMarkers";
import { paragraphNumberingLevel, paragraphNumberingReferenceId } from "../docx/numberingReference";
import type { CounterFormat } from "../types/document";
import {
  effectiveParagraphNumbering,
  paragraphNumberingSourcesFromBoundary,
} from "./numberingAttr";
import type { ParagraphAttrs } from "./schema/nodes";
import { hasListRendering, type ListRenderingAttrKey } from "./listRenderingAttrs";

export type ListCounterState = {
  counters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  seenLevels: Set<string>;
  restartedNumIds?: Set<number>;
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

export const MAX_LIST_LEVEL = 8;

export { LIST_RENDERING_ATTR_KEYS } from "./listRenderingAttrs";

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
  listFoldedMarkerSuffix: null,
  listMarkerSecondSlotOffsetTwips: null,
  listLevelNumFmts: null,
  listLevelStarts: null,
  listLevelTabs: null,
  listAbstractNumId: null,
  listStartOverride: null,
} satisfies Record<ListRenderingAttrKey, null>);

export function createListCounterState(): ListCounterState {
  return {
    counters: new Map(),
    abstractCounters: new Map(),
    seenLevels: new Set(),
    restartedNumIds: new Set(),
    previousList: { abstractNumId: null, fromStyle: false, numId: null },
  };
}

export function cloneListCounterState(state: ListCounterState): ListCounterState {
  const clonedArrays = new Map<number[], number[]>();
  const cloneCounters = (source: number[]): number[] => {
    const existing = clonedArrays.get(source);
    if (existing) return existing;
    const cloned = [...source];
    clonedArrays.set(source, cloned);
    return cloned;
  };
  const counters = new Map<number, number[]>();
  for (const [numId, values] of state.counters) {
    counters.set(numId, cloneCounters(values));
  }
  const abstractCounters = new Map<number, number[]>();
  for (const [abstractNumId, values] of state.abstractCounters) {
    abstractCounters.set(abstractNumId, cloneCounters(values));
  }
  return {
    counters,
    abstractCounters,
    seenLevels: new Set(state.seenLevels),
    restartedNumIds: new Set(state.restartedNumIds ?? []),
    previousList: {
      ...(state.previousList ?? { abstractNumId: null, fromStyle: false, numId: null }),
    },
    ...(state.lastAdvancedNumId !== undefined
      ? { lastAdvancedNumId: state.lastAdvancedNumId }
      : {}),
  };
}

function previousListAttrs(attrs: ParagraphAttrs): ParagraphAttrs | null {
  const change = attrs._propertyChanges?.at(0);
  if (!change) return null;
  const previous = change.previousFormatting ?? {};
  return {
    ...(previous.numPr == null ? {} : { numPr: previous.numPr }),
    ...(previous.numPrFromStyle == null ? {} : { numPrFromStyle: previous.numPrFromStyle }),
    ...(previous.listIsBullet !== undefined ? { listIsBullet: previous.listIsBullet } : {}),
    ...(previous.listIsLegal !== undefined ? { listIsLegal: previous.listIsLegal } : {}),
    ...(previous.listNumFmt !== undefined ? { listNumFmt: previous.listNumFmt } : {}),
    ...(previous.listMarker !== undefined ? { listMarker: previous.listMarker } : {}),
    ...(previous.listMarkerTemplate !== undefined
      ? { listMarkerTemplate: previous.listMarkerTemplate }
      : {}),
    ...(previous.listFoldedMarkerSuffix !== undefined
      ? { listFoldedMarkerSuffix: previous.listFoldedMarkerSuffix }
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
    ...(previous.listLevelTabs !== undefined ? { listLevelTabs: previous.listLevelTabs } : {}),
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
  if (attrs.pPrMark?.kind === "del") {
    const visible = advance(previous ?? attrs, "original");
    return { ...visible, advances };
  }
  const final = advance(attrs, "final");
  const original =
    attrs.pPrMark?.kind === "ins" ? undefined : advance(previous ?? attrs, "original");
  const levelOnlyResolves =
    attrs.numPr?.kind !== "levelOnly" ||
    effectiveParagraphNumbering(
      paragraphNumberingSourcesFromBoundary({
        numPr: attrs.numPr,
        numPrFromStyle: attrs.numPrFromStyle,
      }),
    ).kind === "reference";
  const visible =
    attrs.numPr?.kind !== "none" &&
    levelOnlyResolves &&
    final.marker === null &&
    original &&
    original.marker !== null &&
    previous?.numPr?.kind === "reference"
      ? original
      : final;
  return { ...visible, advances };
}

export function formatCounter(value: number, format: CounterFormat | undefined): string {
  return formatOoxmlCounter(value, format);
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
  if (!hasListRendering(attrs)) {
    state.previousList = { abstractNumId: null, fromStyle: false, numId: null };
    return null;
  }
  const markerTemplate = attrs.listMarkerTemplate ?? attrs.listMarker;
  const foldedMarkerSuffix = attrs.listFoldedMarkerSuffix ?? undefined;
  const effectiveNumbering = effectiveParagraphNumbering(
    paragraphNumberingSourcesFromBoundary({
      numPr: attrs.numPr,
      numPrFromStyle: attrs.numPrFromStyle,
    }),
  );
  const level = paragraphNumberingLevel(effectiveNumbering) ?? 0;
  if (!Number.isInteger(level) || level < 0 || level > MAX_LIST_LEVEL) {
    return null;
  }
  const numId = paragraphNumberingReferenceId(effectiveNumbering);
  if (numId === undefined) {
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
    if (foldedMarkerSuffix !== undefined && attrs.listMarkerTemplate == null) {
      return attrs.listMarker ?? null;
    }
    const marker = convertBulletToUnicode(
      attrs.listMarkerTemplate ?? attrs.listMarker ?? "",
      bulletMarkerFontName(attrs.listMarkerFormatting),
    );
    return foldedMarkerSuffix === undefined ? marker : `${marker}\t${foldedMarkerSuffix}`;
  }

  const firstInstanceEncounter = !state.counters.has(numId);
  let counters = state.counters.get(numId) ?? Array.from({ length: 9 }, () => Number.NaN);
  const abstractNumId = attrs.listAbstractNumId;
  const previousList = state.previousList ?? {
    abstractNumId: null,
    fromStyle: false,
    numId: null,
  };
  const restartedNumIds = (state.restartedNumIds ??= new Set());
  const latestAbstractCounters =
    abstractNumId === undefined ? undefined : state.abstractCounters.get(abstractNumId);
  const styleNumbering =
    attrs.numPr?.kind !== "reference" && attrs.numPrFromStyle?.kind === "reference";
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
  if (
    latestAbstractCounters &&
    (styleNumbering || resumesRestartedInstance || resumesStyleInstance)
  ) {
    counters = latestAbstractCounters;
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
    state.abstractCounters.set(abstractNumId, counters);
  }
  state.previousList = {
    abstractNumId: abstractNumId ?? null,
    fromStyle: Boolean(styleNumbering),
    numId,
  };

  const levelFormats =
    attrs.listLevelNumFmts ?? (attrs.listNumFmt ? [attrs.listNumFmt] : undefined);
  if (
    markerTemplate?.includes("%") &&
    (foldedMarkerSuffix === undefined || attrs.listMarkerTemplate != null)
  ) {
    const marker = resolveListTemplate({
      template: markerTemplate,
      counters,
      levelFormats,
      forceDecimal: attrs.listIsLegal,
    });
    return foldedMarkerSuffix === undefined ? marker : `${marker}\t${foldedMarkerSuffix}`;
  }
  if (foldedMarkerSuffix !== undefined && attrs.listMarkerTemplate != null) {
    return `${attrs.listMarkerTemplate}\t${foldedMarkerSuffix}`;
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
