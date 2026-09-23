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
  /** Stop re-running layout and reporting; pending loads are ignored. */
  dispose: () => void;
};

export type HyphenationReadinessOptions = {
  relayout: () => void;
  onError: (error: HyphenationDictionaryError) => void;
};

export const createHyphenationReadiness = ({
  relayout,
  onError,
}: HyphenationReadinessOptions): HyphenationReadiness => {
  let active = true;
  const pending = new Set<HyphenationDictionaryId>();
  const reported = new Set<HyphenationDictionaryId>();

  // Settles into a relayout or a report, never a rejection:
  // `requestHyphenationDictionary` resolves failures as a settled state.
  const follow = async (dictionary: HyphenationDictionaryId): Promise<void> => {
    pending.add(dictionary);
    const settled = await requestHyphenationDictionary(dictionary);
    pending.delete(dictionary);
    if (!active) {
      return;
    }
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
    dispose: () => {
      active = false;
    },
  };
};
