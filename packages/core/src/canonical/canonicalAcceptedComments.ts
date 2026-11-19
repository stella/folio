import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  findStoryBody,
  type OpStory,
} from "@stll/docx-core/ops";

import { anchoredCommentIdsInBlocks } from "../ai-edits/comment-lifecycle";
import {
  getTrackedChangeGroupsForIds,
  getTrackedChangeStatsFromDoc,
} from "../ai-edits/read";
import { pendingSuggestionRevisionIds } from "../prosemirror/commands/comments";
import { lockedControlSpans } from "../prosemirror/lockedControls";
import type { Document } from "../types/document";
import { canonicalRevisionLocks } from "./canonicalRevisionLocks";

type CanonicalAcceptedCommentsOptions = {
  document: Document;
  projection: PMNode;
  story: OpStory;
};

/** Read accepted anchors with the canonical revision writer. */
export const canonicalAcceptedComments = ({
  document,
  projection,
  story,
}: CanonicalAcceptedCommentsOptions): ReadonlySet<number> | null => {
  const state = EditorState.create({ doc: projection });
  const proposed = pendingSuggestionRevisionIds(state);
  let ids = getTrackedChangeStatsFromDoc(projection).ids.filter(
    (id) => !proposed.has(id),
  );
  if (ids.length > 0 && lockedControlSpans(projection).length > 0) {
    const selected = new Set(ids);
    const groups = getTrackedChangeGroupsForIds(projection, ids).map((group) =>
      group.filter((id) => selected.has(id)),
    );
    ids = canonicalRevisionLocks({
      document,
      story,
      groups,
      decision: "accept",
    }).unlocked.flat();
  }

  let accepted = document;
  if (ids.length > 0) {
    const result = applyDocumentOps(document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds: ids,
        decision: "accept",
      },
    ]);
    if (result.isErr()) return null;
    accepted = result.value.document;
  }
  const body = findStoryBody(accepted, story);
  return body === undefined ? null : anchoredCommentIdsInBlocks(body.content);
};
