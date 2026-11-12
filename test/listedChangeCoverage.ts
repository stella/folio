import type { Node as PMNode } from "prosemirror-model";

import {
  getTrackedChangeGroupsForIds,
  getTrackedChangesFromDoc,
} from "../packages/core/src/ai-edits/read";
import type { Document } from "../packages/core/src/types/document";
import { storyRevisionIds } from "./reviewProjection";

/**
 * How the change list a host reads covers the revisions `document` holds,
 * with `doc` its editor projection. A host resolves by listed id, so every
 * revision must belong to exactly one listed change: one no change reaches
 * stays pending after the host resolves everything it sees, and one two
 * changes share resolves with whichever the host picks first.
 */
export const listedChangeCoverage = (document: Document, doc: PMNode) => {
  const listed = [
    ...new Set(getTrackedChangesFromDoc(doc).map(({ id }) => id)),
  ];
  const groups = listed.map(
    (id) => getTrackedChangeGroupsForIds(doc, [id]).at(0) ?? [id],
  );
  const owners = new Map<number, Set<string>>();
  for (const group of groups) {
    const key = JSON.stringify([...group].sort((left, right) => left - right));
    for (const id of group)
      owners.set(id, (owners.get(id) ?? new Set()).add(key));
  }
  const revisions = storyRevisionIds(document);
  return {
    listed,
    /** Revisions the model holds that no listed change reaches. */
    unreached: revisions.filter((id) => !owners.has(id)),
    /** Revisions more than one listed change reaches. */
    shared: revisions.filter((id) => (owners.get(id)?.size ?? 0) > 1),
  };
};

/** A line naming what the list misses, or null when it covers every revision. */
export const listedChangeCoverageProblem = (
  document: Document,
  doc: PMNode,
): string | null => {
  const { listed, unreached, shared } = listedChangeCoverage(document, doc);
  if (unreached.length === 0 && shared.length === 0) return null;
  return `the change list ${JSON.stringify(listed)} leaves revisions ${JSON.stringify(unreached)} unreached and ${JSON.stringify(shared)} shared`;
};
