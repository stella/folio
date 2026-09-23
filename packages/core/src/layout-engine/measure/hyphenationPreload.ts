/**
 * Awaitable dictionary loading for hosts that measure once and cannot re-run
 * layout when a dictionary arrives (headless layout, tests).
 */

import { Result } from "better-result";

import {
  type HyphenationDictionaryError,
  type HyphenationDictionaryId,
  hyphenationDictionaryFor,
  requestHyphenationDictionary,
} from "./hyphenationDictionaries";

/**
 * Load the dictionaries `locales` need before measuring, so a synchronous
 * layout pass hyphenates on its first run. Locales without a dictionary are
 * ignored.
 */
export const preloadHyphenationDictionaries = async (
  locales: Iterable<string>,
): Promise<Result<void, HyphenationDictionaryError>> => {
  const dictionaries = new Set<HyphenationDictionaryId>();
  for (const locale of locales) {
    const dictionary = hyphenationDictionaryFor(locale);
    if (dictionary !== undefined) {
      dictionaries.add(dictionary);
    }
  }
  const settled = await Promise.all([...dictionaries].map(requestHyphenationDictionary));
  for (const outcome of settled) {
    if (outcome.status === "failed") {
      return Result.err(outcome.error);
    }
  }
  return Result.ok(undefined);
};
