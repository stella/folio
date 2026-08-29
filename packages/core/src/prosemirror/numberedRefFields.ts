import { Fragment, type Node as PMNode } from "prosemirror-model";

import { expectBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import { expectFieldAttrs, expectParagraphAttrs } from "./attrs";
import {
  advanceVisibleListMarker,
  createListCounterState,
  MAX_LIST_LEVEL,
  resolveListTemplateWithComponents,
  type ListCounterState,
  type ListCounterStream,
  type ListCounterStreams,
  type ResolvedListComponent,
} from "./listMarker";

const MAX_FIELD_INSTRUCTION_CHARS = 2048;
const MAX_BOOKMARK_NAME_CHARS = 256;
const MAX_NUMBERED_REF_FIELDS = 1024;
const MAX_BOOKMARK_TARGETS = 2048;

type NumberSwitch = "n" | "r" | "w";

type NumberedRefSpec = {
  bookmark: string;
  numberSwitch: NumberSwitch;
};

type NumberedTarget = {
  current: string;
  full: string | null;
  path: NumberingPath | null;
};

type NumberingScheme = { type: "abstract"; id: number } | { type: "instance"; id: number };

type NumberingPath = {
  scheme: NumberingScheme;
  counters: readonly number[];
  rendered: RenderedNumber;
};

type RenderedNumber = {
  value: string;
  components: readonly ResolvedListComponent[];
};

type NumberingContexts = {
  byNumId: Map<number, (RenderedNumber | undefined)[]>;
  byAbstractNumId: Map<number, (RenderedNumber | undefined)[]>;
};

type NumberingContextStreams = Record<ListCounterStream, NumberingContexts>;

export type NumberedRefResolutionOptions = {
  /** Counter snapshot at the start of this body story. The resolver consumes it in place. */
  listCounterState?: ListCounterState;
  /** Counter snapshot for tracked deletions in the original document view. */
  originalListCounterState?: ListCounterState;
};

function tokenizeInstruction(instruction: string): string[] | null {
  const tokens: string[] = [];
  let index = 0;
  while (index < instruction.length) {
    while (/\s/u.test(instruction[index] ?? "")) {
      index += 1;
    }
    if (index >= instruction.length) {
      break;
    }
    if (instruction[index] === '"') {
      const close = instruction.indexOf('"', index + 1);
      if (close === -1) {
        return null;
      }
      tokens.push(instruction.slice(index + 1, close));
      index = close + 1;
      continue;
    }
    let end = index;
    while (end < instruction.length && !/\s/u.test(instruction[end] ?? "")) {
      end += 1;
    }
    tokens.push(instruction.slice(index, end));
    index = end;
  }
  return tokens;
}

export function parseNumberedRefInstruction(instruction: string): NumberedRefSpec | null {
  if (instruction.length > MAX_FIELD_INSTRUCTION_CHARS) {
    return null;
  }
  const tokens = tokenizeInstruction(instruction);
  if (tokens === null || tokens.length < 3 || tokens.at(0)?.toUpperCase() !== "REF") {
    return null;
  }
  const bookmark = tokens.at(1);
  if (!bookmark || bookmark.length > MAX_BOOKMARK_NAME_CHARS || bookmark.startsWith("\\")) {
    return null;
  }

  let numberSwitch: NumberSwitch | null = null;
  for (let index = 2; index < tokens.length; index += 1) {
    const token = tokens[index]?.toUpperCase();
    if (token === "\\N" || token === "\\R" || token === "\\W") {
      if (numberSwitch !== null) {
        return null;
      }
      if (token === "\\N") {
        numberSwitch = "n";
      } else if (token === "\\R") {
        numberSwitch = "r";
      } else {
        numberSwitch = "w";
      }
      continue;
    }
    if (token === "\\H") {
      continue;
    }
    if (token === "\\*MERGEFORMAT") {
      continue;
    }
    if (token === "\\*" && tokens[index + 1]?.toUpperCase() === "MERGEFORMAT") {
      index += 1;
      continue;
    }
    return null;
  }
  return numberSwitch === null ? null : { bookmark, numberSwitch };
}

function trimTrailingPeriod(value: string): string {
  return value.length > 1 && value.endsWith(".") ? value.slice(0, -1) : value;
}

function firstReferencedLevel(template: string): number | null {
  let first: number | null = null;
  for (const match of template.matchAll(/%(?<level>[1-9])/gu)) {
    const level = Number.parseInt(match.groups?.["level"] ?? "", 10) - 1;
    if (Number.isNaN(level)) {
      continue;
    }
    first = first === null ? level : Math.min(first, level);
  }
  return first;
}

function numberTargetForParagraph(
  node: PMNode,
  contextStreams: NumberingContextStreams,
  streams: ListCounterStreams,
): NumberedTarget | null {
  const resolution = advanceVisibleListMarker(expectParagraphAttrs(node), streams);
  let visibleTarget: NumberedTarget | null = null;
  for (const advanced of resolution.advances) {
    const target = numberTargetForAdvancedMarker(
      advanced.counterAttrs,
      advanced.marker,
      advanced.counterState,
      contextStreams[advanced.stream],
    );
    if (
      advanced.stream === resolution.stream &&
      advanced.counterAttrs === resolution.counterAttrs
    ) {
      visibleTarget = target;
    }
  }
  return visibleTarget;
}

function numberTargetForAdvancedMarker(
  attrs: ReturnType<typeof expectParagraphAttrs>,
  resolvedMarker: string | null,
  state: ListCounterState,
  contexts: NumberingContexts,
): NumberedTarget | null {
  const numId = attrs.numPr?.numId;
  if (
    numId === undefined ||
    numId === 0 ||
    attrs.listIsBullet ||
    attrs.listMarkerHidden ||
    attrs.listNumFmt === "none" ||
    !resolvedMarker
  ) {
    return null;
  }
  const marker = attrs.listMarkerAllCaps ? resolvedMarker.toLocaleUpperCase() : resolvedMarker;

  const level = attrs.numPr?.ilvl ?? 0;
  if (!Number.isInteger(level) || level < 0 || level > MAX_LIST_LEVEL) {
    return null;
  }
  const counters = state.counters.get(numId)?.slice(0, level + 1);
  const levelFormats =
    attrs.listLevelNumFmts ?? (attrs.listNumFmt ? [attrs.listNumFmt] : undefined);
  let renderedMarker: RenderedNumber = { value: marker, components: [] };
  if (attrs.listMarker?.includes("%") && counters?.length === level + 1) {
    const rendered = resolveListTemplateWithComponents({
      template: attrs.listMarker,
      counters,
      levelFormats,
      forceDecimal: attrs.listIsLegal,
    });
    const renderedValue = attrs.listMarkerAllCaps
      ? rendered.value.toLocaleUpperCase()
      : rendered.value;
    if (renderedValue === marker) {
      renderedMarker = { value: marker, components: rendered.components };
    }
  }
  const abstractContexts =
    attrs.listAbstractNumId === undefined
      ? undefined
      : contexts.byAbstractNumId.get(attrs.listAbstractNumId);
  const ownContexts = contexts.byNumId.get(numId) ?? [...(abstractContexts ?? [])];
  const firstLevel = firstReferencedLevel(attrs.listMarker ?? "");
  let renderedFull: RenderedNumber | null;
  if (level === 0 || firstLevel === 0) {
    renderedFull = renderedMarker;
  } else {
    const parentLevel = (firstLevel ?? level) - 1;
    const prefix = ownContexts[parentLevel] ?? abstractContexts?.[parentLevel];
    renderedFull =
      prefix === undefined
        ? null
        : {
            value: `${prefix.value}${marker}`,
            components: [
              ...prefix.components,
              ...renderedMarker.components.map((component) => ({
                level: component.level,
                relativeStart: component.relativeStart + prefix.value.length,
                start: component.start + prefix.value.length,
                end: component.end + prefix.value.length,
              })),
            ],
          };
  }

  ownContexts[level] = renderedFull ?? undefined;
  ownContexts.splice(level + 1);
  contexts.byNumId.set(numId, ownContexts);
  if (attrs.listAbstractNumId !== undefined) {
    contexts.byAbstractNumId.set(attrs.listAbstractNumId, [...ownContexts]);
  }

  const path =
    counters?.length === level + 1 && counters.every(Number.isFinite) && renderedFull !== null
      ? {
          scheme:
            attrs.listAbstractNumId === undefined
              ? { type: "instance" as const, id: numId }
              : { type: "abstract" as const, id: attrs.listAbstractNumId },
          counters,
          rendered: renderedFull,
        }
      : null;

  return {
    current: trimTrailingPeriod(marker),
    full: renderedFull === null ? null : trimTrailingPeriod(renderedFull.value),
    path,
  };
}

function sameNumberingScheme(left: NumberingScheme, right: NumberingScheme): boolean {
  return left.type === right.type && left.id === right.id;
}

function relativeNumber(target: NumberedTarget, source: NumberedTarget | null): string | null {
  if (source === null || source.path === null || target.path === null) {
    return target.full;
  }
  if (!sameNumberingScheme(source.path.scheme, target.path.scheme)) {
    return target.full;
  }

  let sharedLevels = 0;
  const comparedLevels = Math.min(source.path.counters.length, target.path.counters.length);
  while (
    sharedLevels < comparedLevels &&
    source.path.counters[sharedLevels] === target.path.counters[sharedLevels]
  ) {
    sharedLevels += 1;
  }
  if (sharedLevels === 0) {
    return target.full;
  }
  if (sharedLevels >= target.path.counters.length) {
    return null;
  }

  const firstRelativeComponent = target.path.rendered.components.find(
    (component) => component.level >= sharedLevels,
  );
  if (firstRelativeComponent === undefined) {
    return null;
  }
  const suffix = target.path.rendered.value
    .slice(firstRelativeComponent.relativeStart)
    .replace(/^[\s,.:;/-]+/u, "");
  return suffix.length === 0 ? null : trimTrailingPeriod(suffix);
}

function normalizeCachedResult(value: string): string {
  return value.replaceAll("\u00a0", " ").replace(/\s+/gu, " ").trim();
}

type NumberedRefCandidate = {
  node: PMNode;
  cached: string;
  value: string;
};

function collectNumberedRefCandidates(
  doc: PMNode,
  options: NumberedRefResolutionOptions,
): NumberedRefCandidate[] {
  const fields: { node: PMNode; spec: NumberedRefSpec; cached: string }[] = [];
  doc.descendants((node) => {
    if (node.type.name !== "field" && node.type.name !== "structuredField") {
      return true;
    }
    if (fields.length >= MAX_NUMBERED_REF_FIELDS) {
      return false;
    }
    const attrs = expectFieldAttrs(node);
    if (attrs.fieldType !== "REF" || attrs.fldLock) {
      return false;
    }
    const spec = parseNumberedRefInstruction(attrs.instruction);
    if (spec !== null) {
      fields.push({ node, spec, cached: attrs.displayText });
    }
    return false;
  });
  if (fields.length === 0) {
    return [];
  }

  const referencedBookmarks = new Set(fields.map(({ spec }) => spec.bookmark));
  const fieldNodes = new Set(fields.map(({ node }) => node));
  const sourceByField = new WeakMap<PMNode, NumberedTarget | null>();
  const targets = new Map<string, NumberedTarget | null>();
  const streams: ListCounterStreams = {
    final: options.listCounterState ?? createListCounterState(),
    original: options.originalListCounterState ?? createListCounterState(),
  };
  const createContexts = (): NumberingContexts => ({
    byNumId: new Map(),
    byAbstractNumId: new Map(),
  });
  const contextStreams: NumberingContextStreams = {
    final: createContexts(),
    original: createContexts(),
  };

  doc.descendants((node) => {
    if (node.type.name !== "paragraph") {
      return true;
    }

    const target = numberTargetForParagraph(node, contextStreams, streams);
    node.descendants((descendant) => {
      if (descendant.type.name === "bookmarkBoundary") {
        const attrs = expectBookmarkBoundaryAttrs(descendant);
        if (
          attrs.type === "start" &&
          targets.size < MAX_BOOKMARK_TARGETS &&
          referencedBookmarks.has(attrs.name) &&
          !targets.has(attrs.name)
        ) {
          targets.set(attrs.name, target);
        }
        return false;
      }
      if (descendant.type.name !== "field" && descendant.type.name !== "structuredField") {
        return true;
      }
      if (fieldNodes.has(descendant)) {
        sourceByField.set(descendant, target);
      }
      return true;
    });
    const paragraphAttrs = expectParagraphAttrs(node);
    if (targets.size < MAX_BOOKMARK_TARGETS) {
      for (const bookmark of paragraphAttrs.bookmarks ?? []) {
        if (referencedBookmarks.has(bookmark.name) && !targets.has(bookmark.name)) {
          targets.set(bookmark.name, target);
        }
      }
    }
    return false;
  });

  const candidates: NumberedRefCandidate[] = [];
  for (const field of fields) {
    const target = targets.get(field.spec.bookmark);
    if (target === undefined || target === null) {
      continue;
    }
    let value: string | null;
    if (field.spec.numberSwitch === "n") {
      value = target.current;
    } else if (field.spec.numberSwitch === "r") {
      value = relativeNumber(target, sourceByField.get(field.node) ?? null);
    } else {
      value = target.full;
    }
    if (value !== null) {
      candidates.push({ node: field.node, cached: field.cached, value });
    }
  }
  return candidates;
}

const isCalibratedCandidate = ({ node, cached, value }: NumberedRefCandidate): boolean => {
  const normalizedCache = normalizeCachedResult(cached);
  const baseline = expectFieldAttrs(node)._numberedRefBaseline;
  return (
    normalizedCache.length === 0 ||
    normalizedCache === normalizeCachedResult(value) ||
    (baseline !== undefined && normalizedCache === normalizeCachedResult(baseline))
  );
};

const mapDocument = (node: PMNode, transform: (node: PMNode) => PMNode): PMNode => {
  let mapped = node;
  if (node.childCount > 0) {
    const children: PMNode[] = [];
    let changed = false;
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child) => {
      const mappedChild = mapDocument(child, transform);
      children.push(mappedChild);
      changed ||= mappedChild !== child;
    });
    if (changed) {
      mapped = node.copy(Fragment.fromArray(children));
    }
  }
  return transform(mapped);
};

/** Stamp only proven numbered REF fields with clone-safe calibration state. */
export function stampNumberedRefFieldBaselines(doc: PMNode): PMNode {
  const baselines = new WeakMap<PMNode, string>();
  let hasBaselines = false;
  for (const candidate of collectNumberedRefCandidates(doc, {})) {
    if (isCalibratedCandidate(candidate)) {
      baselines.set(candidate.node, candidate.cached);
      hasBaselines = true;
    }
  }
  if (!hasBaselines) {
    return doc;
  }
  return mapDocument(doc, (node) => {
    const baseline = baselines.get(node);
    if (
      baseline === undefined ||
      (node.type.name !== "field" && node.type.name !== "structuredField")
    ) {
      return node;
    }
    return node.type.create(
      { ...node.attrs, _numberedRefBaseline: baseline },
      node.content,
      node.marks,
    );
  });
}

/**
 * Resolve bounded, numbered REF fields against bookmarks in the ProseMirror body story.
 * Unsupported or uncalibrated fields are absent from the returned map and retain their cache.
 */
export function resolveNumberedRefFields(
  doc: PMNode,
  options: NumberedRefResolutionOptions = {},
): ReadonlyMap<PMNode, string> {
  const results = new Map<PMNode, string>();
  for (const candidate of collectNumberedRefCandidates(doc, options)) {
    if (isCalibratedCandidate(candidate)) {
      results.set(candidate.node, candidate.value);
    }
  }
  return results;
}
