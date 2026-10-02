/**
 * Numbering for one story (a header, footer or note) edited in its own editor.
 *
 * Each story editor defines the lists it starts against the package's
 * definitions and its own paragraphs; it cannot see the lists another story
 * started. Two stories can therefore pick the same new `w:numId` for
 * different lists. When a story is written back after another one already
 * defined an id it picked, its paragraphs move to fresh ids first, so each
 * list keeps its own definition.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { Transform } from "prosemirror-transform";

import {
  paragraphNumberingLevel,
  paragraphNumberingReference,
  paragraphNumberingReferenceId,
} from "../docx/numberingReference";
import { getCachedNumberingMap, isBulletLevel } from "../docx/numberingParser";
import { createNumberingIdAllocator } from "../docx/numberingIds";
import type { NumberingDefinitions } from "../types/document";
import { expectParagraphAttrs } from "./attrs";
import { completeNumberingForDoc, paragraphListReferences } from "./listInstanceReferences";
import { paragraphNumberingAttr } from "./numberingAttr";
import { getPackageNumberingDefinitions } from "./plugins/documentNumbering";
import type { ParagraphAttrs } from "./schema/nodes";

type Remap = {
  numIds: Map<number, number>;
  /** Abstract definitions the package already had; a moved list keeps those. */
  baseAbstractNumIds: ReadonlySet<number>;
  abstractNumIds: Map<number, number>;
  abstractIds: ReturnType<typeof createNumberingIdAllocator>;
};

const remappedNumPr = (
  numPr: ParagraphAttrs["numPr"] | null | undefined,
  numIds: ReadonlyMap<number, number>,
): ParagraphAttrs["numPr"] | null | undefined => {
  const numId = paragraphNumberingReferenceId(numPr ?? undefined);
  const next = numId === undefined ? undefined : numIds.get(numId);
  if (next === undefined) {
    return numPr;
  }
  const ilvl = paragraphNumberingLevel(numPr ?? undefined);
  return paragraphNumberingAttr(
    paragraphNumberingReference(ilvl === undefined ? { numId: next } : { numId: next, ilvl }),
  );
};

const remappedAbstractNumId = (
  abstractNumId: number | null | undefined,
  remap: Remap,
): number | null | undefined => {
  if (
    abstractNumId === null ||
    abstractNumId === undefined ||
    remap.baseAbstractNumIds.has(abstractNumId)
  ) {
    return abstractNumId;
  }
  let next = remap.abstractNumIds.get(abstractNumId);
  if (next === undefined) {
    next = remap.abstractIds.next();
    remap.abstractNumIds.set(abstractNumId, next);
  }
  return next;
};

type StoryListNumbering = {
  /** The story with every colliding list moved to its own ids. */
  doc: PMNode;
  /** `numbering` plus the lists the story defines. */
  numbering: NumberingDefinitions | undefined;
};

/**
 * The story in `state` as it should be written into a package whose numbering
 * is now `numbering`, and that numbering completed with the story's lists.
 */
export const storyListNumbering = (
  state: EditorState,
  numbering: NumberingDefinitions | undefined,
): StoryListNumbering => {
  const packageDefinitions = getPackageNumberingDefinitions(state);
  const storyBase = new Set((packageDefinitions?.nums ?? []).map(({ numId }) => numId));
  const defined = new Set((numbering?.nums ?? []).map(({ numId }) => numId));
  const definedMap = numbering ? getCachedNumberingMap(numbering) : null;
  const referenced = new Set<number>();
  const colliding = new Set<number>();
  const existingAbstractNumIds = new Set([
    ...(numbering?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
    ...(packageDefinitions?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
  ]);
  state.doc.descendants((node) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    for (const numId of paragraphListReferences(node)) {
      referenced.add(numId);
    }
    // An id this story defined itself that the package now defines as another
    // list. The story's own earlier save defines it as the same list.
    const attrs = expectParagraphAttrs(node);
    if (typeof attrs.listAbstractNumId === "number")
      existingAbstractNumIds.add(attrs.listAbstractNumId);
    for (const change of attrs._propertyChanges ?? []) {
      const abstractId = change.previousFormatting?.listAbstractNumId;
      if (typeof abstractId === "number") existingAbstractNumIds.add(abstractId);
    }
    const numId = paragraphNumberingReferenceId(attrs.numPr);
    if (numId !== undefined && !storyBase.has(numId) && definedMap?.hasNumbering(numId)) {
      const level = definedMap.getLevel(numId, paragraphNumberingLevel(attrs.numPr) ?? 0);
      const sameList =
        level !== null &&
        isBulletLevel(level) === (attrs.listIsBullet === true) &&
        (attrs.listAbstractNumId ?? undefined) ===
          (definedMap.getAbstractNumId(numId) ?? undefined);
      if (!sameList) {
        colliding.add(numId);
      }
    }
    return false;
  });
  if (colliding.size === 0) {
    return { doc: state.doc, numbering: completeNumberingForDoc(numbering, state.doc) };
  }

  const existingNumIds = [...defined, ...referenced];
  const numIds = createNumberingIdAllocator("num", existingNumIds);
  const remap: Remap = {
    numIds: new Map([...colliding].map((numId) => [numId, numIds.next()])),
    baseAbstractNumIds: new Set(
      (packageDefinitions?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
    ),
    abstractNumIds: new Map(),
    abstractIds: createNumberingIdAllocator("abstract", existingAbstractNumIds),
  };
  const tr = new Transform(state.doc);
  state.doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const attrs = expectParagraphAttrs(node);
    const numId = paragraphNumberingReferenceId(attrs.numPr);
    const moves = numId !== undefined && remap.numIds.has(numId);
    const changes = attrs._propertyChanges?.map((change) => {
      const previous = change.previousFormatting;
      const previousNumId = paragraphNumberingReferenceId(previous?.numPr ?? undefined);
      if (!previous || previousNumId === undefined || !remap.numIds.has(previousNumId)) {
        return change;
      }
      return {
        ...change,
        previousFormatting: {
          ...previous,
          numPr: remappedNumPr(previous.numPr, remap.numIds),
          listAbstractNumId: remappedAbstractNumId(previous.listAbstractNumId, remap) ?? undefined,
        },
      };
    });
    const changesMoved = changes?.some(
      (change, index) => change !== attrs._propertyChanges?.[index],
    );
    if (moves || changesMoved) {
      tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        ...(moves
          ? {
              numPr: remappedNumPr(attrs.numPr, remap.numIds),
              listAbstractNumId: remappedAbstractNumId(attrs.listAbstractNumId, remap) ?? null,
            }
          : {}),
        ...(changesMoved ? { _propertyChanges: changes } : {}),
      });
    }
    return false;
  });
  return { doc: tr.doc, numbering: completeNumberingForDoc(numbering, tr.doc) };
};
