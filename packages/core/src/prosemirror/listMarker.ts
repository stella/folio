import { formatOoxmlCounter } from "../docx/ooxmlCounterFormatter";
import { convertBulletToUnicode } from "../docx/bulletMarkers";
import type { NumberFormat } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

export type ListCounterState = {
  counters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  seenLevels: Set<string>;
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

export function createListCounterState(): ListCounterState {
  return {
    counters: new Map(),
    abstractCounters: new Map(),
    seenLevels: new Set(),
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
    ...(previous.listMarkerHidden !== undefined
      ? { listMarkerHidden: previous.listMarkerHidden }
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
  const level = attrs.numPr?.ilvl ?? 0;
  if (!Number.isInteger(level) || level < 0 || level > MAX_LIST_LEVEL) {
    return null;
  }
  const numId = attrs.numPr?.numId;
  if (numId === undefined || numId === 0) {
    if (attrs.listMarker?.includes("%") && !attrs.listIsBullet) {
      const counters = getLastListCounters(state);
      if (counters) {
        return resolveListTemplate({
          template: attrs.listMarker,
          counters,
          levelFormats: attrs.listLevelNumFmts,
          forceDecimal: attrs.listIsLegal,
        });
      }
    }
    return null;
  }

  if (attrs.listIsBullet) {
    return convertBulletToUnicode(attrs.listMarker ?? "");
  }

  const counters = state.counters.get(numId) ?? Array.from({ length: 9 }, () => Number.NaN);
  const abstractNumId = attrs.listAbstractNumId;
  if (level > 0) {
    const latestAbstractCounters =
      abstractNumId === undefined ? undefined : state.abstractCounters.get(abstractNumId);
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
  if (!state.seenLevels.has(seenKey)) {
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
    state.abstractCounters.set(abstractNumId, [...counters]);
  }

  const levelFormats =
    attrs.listLevelNumFmts ?? (attrs.listNumFmt ? [attrs.listNumFmt] : undefined);
  if (attrs.listMarker?.includes("%")) {
    return resolveListTemplate({
      template: attrs.listMarker,
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
