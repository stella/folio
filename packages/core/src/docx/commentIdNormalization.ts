/**
 * One comment per `w:id`, because that is all the body can address.
 *
 * `w:commentReference`, `w:commentRangeStart` and `w:commentRangeEnd` name a
 * comment by its `w:id`, so two `w:comment` elements sharing an id make every
 * marker that names it ambiguous. Word opens such a package and resolves each
 * marker to the first definition, so folio keeps the first and drops the rest:
 * no marker can address a later one, and re-numbering it would invent a
 * comment nothing anchors. Microsoft's own conformance corpus ships a file
 * that does this.
 *
 * Reply links survive unchanged. `w15:paraIdParent` resolves to a comment id
 * before this runs, and that id still belongs to the definition that stayed.
 */

import type { Comment } from "../types/document";

export type NormalizeCommentIdsResult = {
  droppedDuplicateComments: number;
};

export const normalizeCommentIds = (comments: Comment[]): NormalizeCommentIdsResult => {
  const seen = new Set<number>();
  let kept = 0;
  for (const comment of comments) {
    if (seen.has(comment.id)) {
      continue;
    }
    seen.add(comment.id);
    comments[kept] = comment;
    kept += 1;
  }
  const droppedDuplicateComments = comments.length - kept;
  comments.length = kept;
  return { droppedDuplicateComments };
};
