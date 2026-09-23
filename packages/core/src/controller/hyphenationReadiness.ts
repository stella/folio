/**
 * Per-editor follow-up for hyphenation dictionaries a layout run lacked.
 *
 * Framework-neutral so both adapters share one implementation. The layout
 * pipeline hands {@link HyphenationReadiness.track} the dictionaries its run
 * requested but could not use; this editor then re-runs layout when one loads
 * (the load has already invalidated measured paragraphs) and reports a failed
 * one once. Editors that never asked for a dictionary hear nothing about it,
 * and an editor mounted after a failure still learns of it on its first run.
 */

import {
  type HyphenationDictionaryError,
  type HyphenationDictionaryId,
  requestHyphenationDictionary,
} from "../layout-engine/measure/hyphenationDictionaries";

export type HyphenationReadiness = {
  /** Follow up on the dictionaries one layout run lacked. */
  track: (missing: ReadonlySet<HyphenationDictionaryId>) => void;
  /**
   * Ignore loads still pending, as on unmount. A later `track` follows up
   * afresh, so an adapter whose lifecycle re-attaches (React Strict Mode
   * effects) keeps working.
   */
  cancel: () => void;
};

export type HyphenationReadinessOptions = {
  relayout: () => void;
  onError: (error: HyphenationDictionaryError) => void;
};

export const createHyphenationReadiness = ({
  relayout,
  onError,
}: HyphenationReadinessOptions): HyphenationReadiness => {
  // Bumped by `cancel`; a follow-up from an older epoch settles silently.
  let epoch = 0;
  const pending = new Set<HyphenationDictionaryId>();
  const reported = new Set<HyphenationDictionaryId>();

  // Settles into a relayout or a report, never a rejection:
  // `requestHyphenationDictionary` resolves failures as a settled state.
  const follow = async (dictionary: HyphenationDictionaryId): Promise<void> => {
    const startedIn = epoch;
    pending.add(dictionary);
    const settled = await requestHyphenationDictionary(dictionary);
    if (startedIn !== epoch) {
      return;
    }
    pending.delete(dictionary);
    switch (settled.status) {
      case "loaded":
        relayout();
        return;
      case "failed":
        if (!reported.has(dictionary)) {
          reported.add(dictionary);
          onError(settled.error);
        }
        return;
      default:
        settled satisfies never;
    }
  };

  return {
    track: (missing) => {
      for (const dictionary of missing) {
        if (!pending.has(dictionary)) {
          void follow(dictionary);
        }
      }
    },
    cancel: () => {
      epoch += 1;
      pending.clear();
    },
  };
};
