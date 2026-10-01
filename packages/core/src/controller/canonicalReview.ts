import { getTrackedChangesFromDoc, getTrackedChangeGroupIdsFromDoc } from "../ai-edits/read";
import { findAIEditRevisionRange } from "../prosemirror/commands/comments";
import type { FolioEditor } from "./folioEditor";

type ResolveCanonicalReviewRangeOptions = {
  editor:
    | Pick<FolioEditor, "getCanonicalDocument" | "getState" | "resolveCanonicalRevisions">
    | null
    | undefined;
  from: number;
  to: number;
  resolution: "accept" | "reject";
};

/** Null delegates to the default session's ProseMirror review commands. */
export const resolveCanonicalReviewRange = ({
  editor,
  from,
  to,
  resolution,
}: ResolveCanonicalReviewRangeOptions): boolean | null => {
  if (!editor?.getCanonicalDocument()) return null;
  const state = editor.getState();
  if (!state) return false;
  const ids = new Set<number>();
  for (const change of getTrackedChangesFromDoc(state.doc)) {
    const groupIds = getTrackedChangeGroupIdsFromDoc(state.doc, change.id);
    const range = findAIEditRevisionRange(state, groupIds);
    if (!range) continue;
    const intersects =
      from === to ? range.from <= from && range.to >= to : range.from < to && range.to > from;
    if (intersects) for (const id of groupIds) ids.add(id);
  }
  return ids.size > 0 && editor.resolveCanonicalRevisions([...ids], resolution);
};
