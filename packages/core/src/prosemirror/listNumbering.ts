/**
 * Which numbering instance a list command puts a paragraph in, and the
 * *Restart Numbering*, *Continue Numbering* and *Set Numbering Value* commands.
 *
 * A paragraph made into a list item continues the list directly above or below
 * it when that list is of the requested kind (and, for a typed marker, has the
 * typed format). Otherwise it starts a new list: a new `w:num`, defined on the
 * paragraphs it numbers (see `docx/listNumberingInstances.ts`). A list that a
 * paragraph style supplies (a heading's numbering from `styles.xml`) is never
 * joined implicitly; it belongs to the style, not to the body text around it.
 *
 * Every change made while suggesting is recorded as one `w:pPrChange` per
 * paragraph. A paragraph that already carries one keeps a single record whose
 * previous state is the paragraph's original, so rejecting it restores what the
 * paragraph was before any of the tracked changes.
 */

import type { Node as PMNode, ResolvedPos } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";

import {
  mintListInstance,
  restartListInstance,
  type ListKind,
  type ListLevelFormat,
} from "../docx/listNumberingInstances";
import { createNumberingMap, isBulletLevel, type NumberingMap } from "../docx/numberingParser";
import { paragraphNumberingLevel, paragraphNumberingReferenceId } from "../docx/numberingReference";
import type { ListLevel } from "../types/document";
import { expectParagraphAttrs } from "./attrs";
import { paragraphPropertiesSnapshot } from "./commands/propertyChangeScope";
import { LIST_RENDERING_ATTR_KEYS } from "./listMarker";
import { getDocumentStyleResolver } from "./plugins/documentStyleState";
import type { RunStyleResolver } from "./runStyleFormatting";
import { getDocumentNumbering } from "./plugins/documentNumbering";
import { makeRevisionInfo, SUGGESTION_META } from "./plugins/suggestionMode";
import type { ParagraphAttrs, ParagraphPropertyChangeAttrs } from "./schema/nodes";
import { directParagraphIndentation } from "./paragraphIndentation";
import { listAttrsFromNumbering, listLevelIndentAttrPatch } from "./styles/resolvedStyleAttrs";
import { listRenderingFor, recordsListRendering } from "./listRendering";

const PARAGRAPH_NODE = "paragraph";

/** What a list command or typed marker asks for. */
export type ListRequest = {
  kind: ListKind;
  /**
   * The first level's format. A neighbouring list is continued only when its
   * level has this format; a new list is defined with it.
   */
  format?: ListLevelFormat | undefined;
  /** The value a new list's first item shows. */
  start?: number | undefined;
};

type RevisionInfo = { id: number; author: string; date: string };

// ============================================================================
// TRACKED PARAGRAPH-PROPERTY CHANGES
// ============================================================================

type PreviousFormatting = NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]>;

/**
 * The paragraph's properties as a `w:pPrChange` records them before a list
 * change: every non-null in-scope attr (a reject restores the scope
 * wholesale, so an unrecorded one would be wiped), plus the numbering and the
 * list-rendering bookkeeping with explicit nulls (those sit outside the
 * wholesale scope, so only recorded keys restore).
 */
const listChangeSnapshot = (
  node: PMNode,
  attrs: Record<string, unknown>,
): Record<string, unknown> => {
  const previousFormatting: Record<string, unknown> = paragraphPropertiesSnapshot(
    node.type.create(attrs, node.content, node.marks),
  );
  previousFormatting["numPr"] ??= null;
  for (const key of LIST_RENDERING_ATTR_KEYS) {
    previousFormatting[key] = attrs[key] ?? null;
  }
  return previousFormatting;
};

/**
 * The original state a pending record describes, completed with the list
 * attrs it may not state. A record written by a list command states them all.
 * One written for another property (or read from a file) states the scope
 * only: absent direct numbering uncovers the original style's numbering,
 * so the rendering is recomputed from it rather than copied from
 * the live paragraph, whose list may since have changed.
 */
type OriginalListFormattingOptions = {
  record: PreviousFormatting;
  numbering: NumberingMap | null;
  styleResolver: RunStyleResolver | null;
};

const originalListFormatting = ({
  record,
  numbering,
  styleResolver,
}: OriginalListFormattingOptions): Record<string, unknown> => {
  const original: Record<string, unknown> = { ...record };
  if (recordsListRendering(record)) {
    return original;
  }
  const inherited = styleResolver?.resolveParagraphStyle(record.styleId ?? undefined)
    .paragraphFormatting?.numPr;
  Object.assign(original, listRenderingFor(record.numPr ?? inherited, numbering));
  original["numPr"] = record.numPr ?? null;
  return original;
};

const canonicalJson = (value: unknown): string =>
  JSON.stringify(value, (_key, nested: unknown) =>
    nested !== null && typeof nested === "object" && !Array.isArray(nested)
      ? Object.fromEntries(
          Object.entries(nested)
            .filter(([, entry]) => entry !== null && entry !== undefined)
            .toSorted(([left], [right]) => left.localeCompare(right)),
        )
      : nested,
  );

type TrackListChangeOptions = {
  current: PMNode;
  next: Record<string, unknown>;
  rev: RevisionInfo;
  numbering: NumberingMap | null;
  styleResolver: RunStyleResolver | null;
};

/**
 * `next` with the list change recorded as a tracked paragraph-property change.
 *
 * A paragraph holds at most one serializable `w:pPrChange`. When it already
 * has one, the record is replaced by one that keeps the original previous
 * state: a second change under tracking is a further change to the same
 * paragraph, and rejecting it must restore the paragraph as it was before
 * either. When the change brings the paragraph back to that original, there
 * is nothing left to track and the author's own record is dropped.
 *
 * A pending record by another author (another reviewer's, or one read from
 * the file) keeps its author, id and date: the change joins that revision
 * rather than re-attributing it, and it is never dropped.
 */
const trackListChange = ({
  current,
  next,
  rev,
  numbering,
  styleResolver,
}: TrackListChangeOptions): Record<string, unknown> => {
  const currentAttrs = expectParagraphAttrs(current);
  const existing = currentAttrs._propertyChanges ?? [];
  const pending = existing.find(({ info }) => info.provenance !== "suggested");
  const retained = existing.filter((change) => change !== pending);
  const previousFormatting = pending?.previousFormatting
    ? originalListFormatting({ record: pending.previousFormatting, numbering, styleResolver })
    : listChangeSnapshot(current, currentAttrs);
  const ownPending = pending !== undefined && pending.info.author === rev.author;

  if (
    ownPending &&
    canonicalJson(previousFormatting) === canonicalJson(listChangeSnapshot(current, next))
  ) {
    return { ...next, _propertyChanges: retained.length > 0 ? retained : null };
  }
  const record: ParagraphPropertyChangeAttrs =
    pending && !ownPending
      ? { ...pending, previousFormatting }
      : {
          type: "paragraphPropertyChange",
          info: { id: rev.id, author: rev.author, date: rev.date },
          previousFormatting,
        };
  return { ...next, _propertyChanges: [...retained, record] };
};

// ============================================================================
// WHICH LIST A PARAGRAPH JOINS
// ============================================================================

const levelKind = (level: ListLevel): ListKind => (isBulletLevel(level) ? "bullet" : "numbered");

/** A level's text with its placeholders blanked: `%1.` and `%3.` read alike. */
const levelTextShape = (lvlText: string): string => lvlText.replaceAll(/%\d/gu, "%");

const levelMatches = (level: ListLevel, request: ListRequest): boolean => {
  if (levelKind(level) !== request.kind) {
    return false;
  }
  const { format } = request;
  return (
    format === undefined ||
    (level.numFmt === format.numFmt &&
      levelTextShape(level.lvlText) === levelTextShape(format.lvlText))
  );
};

type ListMembership = {
  numId: number;
  ilvl: number;
  level: ListLevel;
};

/**
 * The list `node` is an item of, when the paragraph states it directly. A
 * numbering its style supplies is the style's, and is not offered to the
 * paragraphs around it.
 */
const directListMembership = (
  node: PMNode | null | undefined,
  numbering: NumberingMap | null,
): ListMembership | null => {
  if (!node || node.type.name !== PARAGRAPH_NODE) {
    return null;
  }
  const attrs = expectParagraphAttrs(node);
  const numId = paragraphNumberingReferenceId(attrs.numPr);
  if (numId === undefined || numId === paragraphNumberingReferenceId(attrs.numPrFromStyle)) {
    return null;
  }
  const ilvl = paragraphNumberingLevel(attrs.numPr) ?? 0;
  const level = numbering?.getLevel(numId, ilvl) ?? null;
  return level ? { numId, ilvl, level } : null;
};

/** The block directly before or after the paragraph `$pos` sits in. */
const siblingOf = ($pos: ResolvedPos, direction: "before" | "after"): PMNode | null => {
  const depth = $pos.depth;
  if (depth === 0) {
    return null;
  }
  const index = $pos.index(depth - 1);
  return $pos.node(depth - 1).maybeChild(direction === "before" ? index - 1 : index + 1);
};

/** Where the paragraphs a command numbers go. */
type ListTarget = {
  numId: number;
  /** The level a paragraph that is not yet a list item takes. */
  ilvl: number;
  /** Definitions the target is resolved in, including one minted for it. */
  numbering: NumberingMap;
};

type ResolveListTargetOptions = {
  numbering: NumberingMap | null;
  $from: ResolvedPos;
  $to: ResolvedPos;
  /** Candidate requests, most likely first (see `listRequestsForMarker`). */
  requests: readonly ListRequest[];
};

/** Every instance a paragraph of `doc` names: directly, through its style, or as a tracked change's previous state. */
const referencedNumIds = (doc: PMNode): Set<number> => {
  const referenced = new Set<number>();
  const add = (numPr: ParagraphAttrs["numPr"] | null | undefined): void => {
    const numId = paragraphNumberingReferenceId(numPr ?? undefined);
    if (numId !== undefined) {
      referenced.add(numId);
    }
  };
  doc.descendants((node) => {
    if (node.type.name !== PARAGRAPH_NODE) {
      return true;
    }
    const attrs = expectParagraphAttrs(node);
    add(attrs.numPr);
    add(attrs.numPrFromStyle);
    for (const change of attrs._propertyChanges ?? []) {
      add(change.previousFormatting?.numPr);
    }
    return false;
  });
  return referenced;
};

/**
 * An instance the package defines for `request` that nothing uses yet: a host
 * that prepares a document with an empty list of each kind means the first
 * list of that kind to be it. An instance tied to a paragraph style (a level
 * naming a `w:pStyle`, or a style link) belongs to that style and is never
 * offered.
 */
const unusedInstanceFor = (
  numbering: NumberingMap,
  doc: PMNode,
  request: ListRequest,
): number | undefined => {
  const referenced = referencedNumIds(doc);
  return numbering.definitions.nums.find(({ numId, abstractNumId }) => {
    const abstract = numbering.getAbstract(abstractNumId);
    const level = numbering.getLevel(numId, 0);
    return (
      !referenced.has(numId) &&
      abstract !== null &&
      abstract.styleLink === undefined &&
      abstract.numStyleLink === undefined &&
      abstract.levels.every(({ pStyle }) => pStyle === undefined) &&
      level !== null &&
      levelMatches(level, request) &&
      (level.start ?? 1) === (request.start ?? level.start ?? 1)
    );
  })?.numId;
};

/**
 * The list a paragraph range joins: the list directly above it, else the one
 * directly below it, when it matches one of `requests`; otherwise a new list:
 * an unused instance the package already defines for the first request, or
 * one defined for it.
 */
export const resolveListTarget = ({
  numbering,
  $from,
  $to,
  requests,
}: ResolveListTargetOptions): ListTarget | null => {
  const [first] = requests;
  if (!first) {
    return null;
  }
  for (const neighbour of [siblingOf($from, "before"), siblingOf($to, "after")]) {
    const membership = directListMembership(neighbour, numbering);
    if (
      numbering &&
      membership &&
      requests.some((request) => levelMatches(membership.level, request))
    ) {
      return { numId: membership.numId, ilvl: membership.ilvl, numbering };
    }
  }
  const unused = numbering ? unusedInstanceFor(numbering, $from.doc, first) : undefined;
  if (numbering && unused !== undefined) {
    return { numId: unused, ilvl: 0, numbering };
  }
  const minted = mintListInstance(numbering?.definitions, first);
  return { numId: minted.numId, ilvl: 0, numbering: createNumberingMap(minted.definitions) };
};

/** A paragraph's attrs as an item of `numId` at `ilvl`, resolved in `numbering`. */
export const listItemAttrs = (
  attrs: Readonly<ParagraphAttrs>,
  { numId, ilvl }: { numId: number; ilvl: number },
  numbering: NumberingMap,
): Record<string, unknown> => ({
  ...attrs,
  ...listAttrsFromNumbering({ numId, ilvl }, numbering),
  ...listLevelIndentAttrPatch(directParagraphIndentation(attrs), { numId, ilvl }, numbering),
  // Counted from the paragraph's own inline fields; the list does not change them.
  listImplicitChildLevelAdvances: attrs["listImplicitChildLevelAdvances"] ?? null,
  listMarkerSecondSlotOffsetTwips: attrs["listMarkerSecondSlotOffsetTwips"] ?? null,
});

type ParagraphUpdate = {
  pos: number;
  node: PMNode;
  next: Record<string, unknown>;
};

type ApplyParagraphUpdatesOptions = {
  tr: Transaction;
  state: EditorState;
  updates: readonly ParagraphUpdate[];
};

/**
 * Write `updates` into `tr`, each recorded as a tracked change while the
 * editor is suggesting. `pos` is a position in `tr.doc`.
 */
export const applyParagraphUpdates = ({
  tr,
  state,
  updates,
}: ApplyParagraphUpdatesOptions): void => {
  const rev = makeRevisionInfo(state);
  const numbering = getDocumentNumbering(state);
  const styleResolver = getDocumentStyleResolver(state);
  for (const { pos, node, next } of updates) {
    const attrs = rev
      ? trackListChange({
          current: node,
          next,
          rev,
          numbering,
          styleResolver,
        })
      : next;
    tr.setNodeMarkup(pos, undefined, attrs);
  }
  if (rev) {
    tr.setMeta(SUGGESTION_META, true);
  }
};

// ============================================================================
// RESTART / CONTINUE / SET VALUE
// ============================================================================

type ListParagraph = { pos: number; node: PMNode; membership: ListMembership };

/** The paragraph the selection starts in, when it is a direct list item. */
const selectedListParagraph = (state: EditorState): ListParagraph | null => {
  const { $from } = state.selection;
  if ($from.parent.type.name !== PARAGRAPH_NODE || $from.depth === 0) {
    return null;
  }
  const membership = directListMembership($from.parent, getDocumentNumbering(state));
  return membership ? { pos: $from.before(), node: $from.parent, membership } : null;
};

/** Every direct item of `numId` from `fromPos` to the end of the document. */
const itemsFrom = (state: EditorState, numId: number, fromPos: number): ListParagraph[] => {
  const numbering = getDocumentNumbering(state);
  const items: ListParagraph[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name !== PARAGRAPH_NODE) {
      return true;
    }
    if (pos >= fromPos) {
      const membership = directListMembership(node, numbering);
      if (membership?.numId === numId) {
        items.push({ pos, node, membership });
      }
    }
    return false;
  });
  return items;
};

type MoveItemsOptions = {
  state: EditorState;
  items: readonly ListParagraph[];
  numId: number;
  numbering: NumberingMap;
};

const moveItems = ({ state, items, numId, numbering }: MoveItemsOptions): Transaction => {
  const tr = state.tr;
  applyParagraphUpdates({
    tr,
    state,
    updates: items.map(({ pos, node, membership }) => ({
      pos,
      node,
      next: listItemAttrs(expectParagraphAttrs(node), { numId, ilvl: membership.ilvl }, numbering),
    })),
  });
  return tr.scrollIntoView();
};

/**
 * Start the list over at `value` from the selected item: *Set Numbering
 * Value* with *Start new list*. The item and every later item of
 * its list move to a new instance of the same definition, whose
 * `w:startOverride` at the item's level is `value`.
 */
export const setNumberingValue =
  (value: number): Command =>
  (state, dispatch) => {
    if (!Number.isSafeInteger(value) || value < 0) {
      return false;
    }
    const selected = selectedListParagraph(state);
    const numbering = getDocumentNumbering(state);
    const abstractNumId = selected ? numbering?.getAbstractNumId(selected.membership.numId) : null;
    if (!selected || !numbering || abstractNumId === null || abstractNumId === undefined) {
      return false;
    }
    if (!dispatch) {
      return true;
    }
    const restarted = restartListInstance(numbering.definitions, {
      abstractNumId,
      ilvl: selected.membership.ilvl,
      start: value,
    });
    dispatch(
      moveItems({
        state,
        items: itemsFrom(state, selected.membership.numId, selected.pos),
        numId: restarted.numId,
        numbering: createNumberingMap(restarted.definitions),
      }),
    );
    return true;
  };

/** *Restart at 1*: the selected item starts its list over at one. */
export const restartNumbering: Command = (state, dispatch) => setNumberingValue(1)(state, dispatch);

/** The nearest list before `beforePos` of the kind `membership` has, other than its own. */
const previousListOfKind = (
  state: EditorState,
  beforePos: number,
  membership: ListMembership,
): number | null => {
  const numbering = getDocumentNumbering(state);
  const kind = levelKind(membership.level);
  let found: number | null = null;
  state.doc.descendants((node, pos) => {
    if (pos >= beforePos) {
      return false;
    }
    if (node.type.name !== PARAGRAPH_NODE) {
      return true;
    }
    const candidate = directListMembership(node, numbering);
    if (candidate && candidate.numId !== membership.numId && levelKind(candidate.level) === kind) {
      found = candidate.numId;
    }
    return false;
  });
  return found;
};

/**
 * *Continue Numbering*: the selected item and the rest of its list
 * join the nearest earlier list of the same kind, so their numbers carry on
 * from it. Not applicable when no such list precedes the item.
 */
export const continueNumbering: Command = (state, dispatch) => {
  const selected = selectedListParagraph(state);
  const numbering = getDocumentNumbering(state);
  if (!selected || !numbering) {
    return false;
  }
  const previous = previousListOfKind(state, selected.pos, selected.membership);
  if (previous === null) {
    return false;
  }
  if (!dispatch) {
    return true;
  }
  dispatch(
    moveItems({
      state,
      items: itemsFrom(state, selected.membership.numId, selected.pos),
      numId: previous,
      numbering,
    }),
  );
  return true;
};

/** What a context menu offers for the item the selection is in. */
export type ListNumberingMenuState =
  | { readonly type: "none" }
  | { readonly type: "listItem"; readonly canContinue: boolean };

export const NO_LIST_NUMBERING_MENU = { type: "none" } as const satisfies ListNumberingMenuState;

/**
 * Whether the selection is in a list item whose numbering can be restarted
 * or set, and whether an earlier list lets it continue: the enabled state of
 * those menu entries.
 */
export const listNumberingMenuState = (state: EditorState): ListNumberingMenuState =>
  restartNumbering(state)
    ? { type: "listItem", canContinue: continueNumbering(state) }
    : NO_LIST_NUMBERING_MENU;
