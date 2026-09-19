/**
 * The one rule for "which `w14:paraId` a comment is threaded by".
 *
 * `w15:commentEx/@w15:paraId` and `w15:paraIdParent` name a comment through the
 * `w14:paraId` of its LAST paragraph, never through its `w:id`. Files in the
 * wild leave that paragraph's id off and carry one on an earlier paragraph, so
 * the rule folio reads by is the last paragraph that carries an id.
 *
 * The rule has two readers over two representations: the parser applies it to
 * the `w:p` children of a `w:comment`, the serializer to the model's
 * `Comment.content`. They must not drift — a save that keyed on a different
 * paragraph than the parse would write a `commentsExtended.xml` whose entries
 * belong to other comments — so both call this instead of restating it.
 */
export const commentThreadParaId = (
  paragraphParaIds: Iterable<string | null | undefined>,
): string | undefined => {
  let key: string | undefined;
  for (const paraId of paragraphParaIds) {
    if (paraId) {
      key = paraId;
    }
  }
  return key;
};
