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

/**
 * Each name's block id as `reviewer` reads now: the first block reading
 * exactly its words, or `undefined` where no words are given or no block
 * reads them.
 */
export const heldBlockIds = <Name extends string>(
  reviewer: FolioDocxReviewer,
  words: Readonly<Record<Name, string | undefined>>,
): ((name: Name) => string | undefined) => {
  const blocks = reviewer.snapshot().blocks;
  return (name) => {
    const text = words[name];
    return text === undefined
      ? undefined
      : blocks.find((block) => block.text === text)?.id;
  };
};
