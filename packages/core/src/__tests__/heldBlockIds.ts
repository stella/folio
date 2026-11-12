/**
 * The ids of blocks a test names by their words, read from a reviewer now.
 *
 * A block's id is its paragraph's `w14:paraId`. Accepting or rejecting a
 * change that joins a paragraph into the next one leaves the paragraph whose
 * mark stays, the next one, under its own id, whatever the first still holds
 * (the reference rule `revisionResolutionReference.test.ts` pins). So an id
 * read before a resolution can name a paragraph that is gone: read the ids
 * again after every accept or reject. The lint rule
 * `folio-held-block-ids/reread-after-resolution` reports an id held across
 * one.
 */

import type { FolioDocxReviewer } from "../ai-edits/headless";

/** Each name's block id, as `heldBlockIds` reads it. */
export type HeldBlockIds<Name extends string> = (
  name: Name,
) => string | undefined;

/**
 * Each name's block id as `reviewer` reads now: the id `previous` held while
 * the reviewer still lists it, else the first block reading exactly the
 * name's words (a block pending deletion reads none), or `undefined` where no
 * words are given or no block reads them.
 */
export const heldBlockIds = <Name extends string>(
  reviewer: FolioDocxReviewer,
  words: Readonly<Record<Name, string | undefined>>,
  previous?: HeldBlockIds<Name>,
): HeldBlockIds<Name> => {
  const blocks = reviewer.snapshot().blocks;
  const listed = new Set(blocks.map((block) => block.id));
  return (name) => {
    const kept = previous?.(name);
    if (kept !== undefined && listed.has(kept)) return kept;
    const text = words[name];
    return text === undefined
      ? undefined
      : blocks.find((block) => block.text === text)?.id;
  };
};
