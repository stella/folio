/**
 * Numbering a session added and nothing uses any more.
 *
 * The canonical session defines a list when a command starts one, while the
 * legacy editor derives definitions from the paragraphs that use them when it
 * saves. Once a rejection or a later edit leaves such a list unused, the
 * session would still write its definition. The save drops the instances and
 * abstract definitions the source package did not have and no paragraph,
 * style or recorded previous formatting names; every definition the source
 * had is kept as it was, used or not.
 */

import type { Document, NumberingDefinitions } from "../types/document";
import { getDocumentSourceNumberingIds } from "./headerFooterVerbatim";

const SKIPPED_PACKAGE_KEYS = new Set(["numbering", "media", "fonts"]);

/** Every `numId` a numbering reference outside the numbering part names. */
const referencedNumIds = (document: Document): Set<number> => {
  const referenced = new Set<number>();
  const visited = new WeakSet<object>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) return;
    if (visited.has(value)) return;
    visited.add(value);
    if (value instanceof Map) {
      for (const entry of value.values()) visit(entry);
      return;
    }
    const numId: unknown = Reflect.get(value, "numId");
    if (Reflect.get(value, "kind") === "reference" && typeof numId === "number")
      referenced.add(numId);
    for (const child of Object.values(value)) visit(child);
  };
  for (const [key, value] of Object.entries(document.package))
    if (!SKIPPED_PACKAGE_KEYS.has(key)) visit(value);
  return referenced;
};

export const withoutUnusedSessionLists = (document: Document): Document => {
  const numbering = document.package.numbering;
  const source = getDocumentSourceNumberingIds(document);
  if (numbering === undefined || source === undefined) return document;
  const referenced = referencedNumIds(document);
  const nums = numbering.nums.filter(
    ({ numId }) => source.nums.has(numId) || referenced.has(numId),
  );
  const usedAbstracts = new Set(nums.map(({ abstractNumId }) => abstractNumId));
  const abstractNums = numbering.abstractNums.filter(
    ({ abstractNumId }) =>
      source.abstractNums.has(abstractNumId) ||
      usedAbstracts.has(abstractNumId),
  );
  if (
    nums.length === numbering.nums.length &&
    abstractNums.length === numbering.abstractNums.length
  )
    return document;
  const kept: NumberingDefinitions = { ...numbering, nums, abstractNums };
  return {
    ...document,
    package: { ...document.package, numbering: kept },
  };
};
