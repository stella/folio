import type { EditorState } from "prosemirror-state";

import { OP_STORIES, type OpStory } from "@stll/docx-core/ops";

import {
  getTrackedChangeGroupIdsFromDoc,
  getTrackedChangeGroupsForIds,
  getTrackedChangesFromDoc,
} from "../ai-edits/read";
import { hasLockedControls } from "../canonical/canonicalLockGuard";
import { canonicalRevisionLocks } from "../canonical/canonicalRevisionLocks";
import { findAIEditRevisionRange } from "../prosemirror/commands/comments";
import { readModel } from "../types/documentModel";
import type { FolioEditor } from "./folioEditor";

export type ReviewRange = { from: number; to: number };

const revisionIdsAtRange = (
  state: EditorState,
  { from, to }: ReviewRange,
): number[] => {
  const ids = new Set<number>();
  for (const change of getTrackedChangesFromDoc(state.doc)) {
    const groupIds = getTrackedChangeGroupIdsFromDoc(state.doc, change.id);
    const range = findAIEditRevisionRange(state, groupIds);
    if (!range) continue;
    const intersects =
      from === to
        ? range.from <= from && range.to >= to
        : range.from < to && range.to > from;
    if (intersects) for (const id of groupIds) ids.add(id);
  }
  return [...ids];
};

export type CanonicalReviewEditor = Pick<
  FolioEditor,
  "getCanonicalStoryProjection" | "getDocument"
>;

type CanonicalReviewLockOptions = {
  editor: CanonicalReviewEditor | null | undefined;
  state: EditorState;
  story: OpStory;
  target:
    | ({ type: "range" } & ReviewRange)
    | { type: "revisions"; ids: readonly number[] };
};

type HeldReviewRevisionIds = Record<"accept" | "reject", ReadonlySet<number>>;

const NO_HELD_REVISIONS: HeldReviewRevisionIds = {
  accept: new Set(),
  reject: new Set(),
};

/** Predict the session's lock decision for the groups its review will settle. */
export const canonicalReviewHeldRevisionIds = ({
  editor,
  state,
  story,
  target,
}: CanonicalReviewLockOptions): HeldReviewRevisionIds => {
  if (!editor?.getCanonicalStoryProjection(OP_STORIES.MAIN))
    return NO_HELD_REVISIONS;
  const model = editor.getDocument();
  if (!model) return NO_HELD_REVISIONS;
  const document = readModel(model);
  if (!hasLockedControls(document)) return NO_HELD_REVISIONS;
  const ids =
    target.type === "range" ? revisionIdsAtRange(state, target) : target.ids;
  const groups = getTrackedChangeGroupsForIds(state.doc, ids);
  const heldIds = (decision: "accept" | "reject") =>
    new Set(
      canonicalRevisionLocks({
        document,
        story,
        groups,
        decision,
      }).held.flatMap((group) => group.ids),
    );
  return { accept: heldIds("accept"), reject: heldIds("reject") };
};

type ResolveCanonicalReviewRangeOptions = {
  editor:
    | Pick<
        FolioEditor,
        "getCanonicalDocument" | "getState" | "resolveCanonicalRevisions"
      >
    | null
    | undefined;
  from: number;
  to: number;
  resolution: "accept" | "reject";
};

/** Null when the editor holds no canonical session, which leaves the review as it is. */
export const resolveCanonicalReviewRange = ({
  editor,
  from,
  to,
  resolution,
}: ResolveCanonicalReviewRangeOptions): boolean | null => {
  if (!editor?.getCanonicalDocument()) return null;
  const state = editor.getState();
  if (!state) return false;
  const ids = revisionIdsAtRange(state, { from, to });
  return ids.length > 0 && editor.resolveCanonicalRevisions(ids, resolution);
};
