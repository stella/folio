/**
 * On-demand hyphenation dictionaries for automatic hyphenation.
 *
 * Each `hyphen` pattern set is tens to hundreds of kilobytes of source whose
 * evaluation builds a pattern trie, and most documents never enable
 * `w:autoHyphenation`. A dictionary is therefore imported only once measurement
 * asks to hyphenate a word in its language. Measurement is synchronous, so the
 * first request starts the import and hyphenates nothing. A layout run wrapped
 * in {@link collectRequestedHyphenationDictionaries} learns which dictionaries
 * it lacked, so its own editor can re-run layout on load or report the failure
 * (`controller/hyphenationReadiness`). Hosts that must lay out correctly on the
 * first pass (headless layout, tests) await `preloadHyphenationDictionaries`
 * (`./hyphenationPreload`) before measuring.
 *
 * This module sits on every editor's import graph, so it speaks in its own
 * settled-state union rather than `Result`, keeping the `Result` API out of
 * adapters' type checks; the preload boundary converts.
 */

import { TaggedError } from "better-result";

import { recordHyphenationDictionaryError } from "../layoutInstrumentation";

export type HyphenateWord = (text: string) => string;

const HYPHENATION_DICTIONARY_IDS = ["cs", "en-gb", "en-us", "sk"] as const;

export type HyphenationDictionaryId = (typeof HYPHENATION_DICTIONARY_IDS)[number];

export type HyphenationDictionaryLoaders = Readonly<
  Record<HyphenationDictionaryId, () => Promise<HyphenateWord>>
>;

// Literal specifiers, one per dictionary, so bundlers emit one lazy chunk each.
const HYPHENATION_DICTIONARY_LOADERS = {
  cs: async () => (await import("hyphen/cs")).default.hyphenateSync,
  "en-gb": async () => (await import("hyphen/en-gb")).default.hyphenateSync,
  "en-us": async () => (await import("hyphen/en-us")).default.hyphenateSync,
  sk: async () => (await import("hyphen/sk")).default.hyphenateSync,
} as const satisfies HyphenationDictionaryLoaders;

let dictionaryLoaders: HyphenationDictionaryLoaders = HYPHENATION_DICTIONARY_LOADERS;

export class HyphenationDictionaryError extends TaggedError("HyphenationDictionaryError")<{
  message: string;
  dictionary: HyphenationDictionaryId;
  cause: unknown;
}> {}

export type SettledHyphenationDictionary =
  | Readonly<{ status: "loaded"; hyphenate: HyphenateWord }>
  // Terminal for the session: retrying on every layout would turn a missing
  // chunk into an endless request/relayout loop.
  | Readonly<{ status: "failed"; error: HyphenationDictionaryError }>;

type HyphenationDictionaryState =
  | Readonly<{ status: "unloaded" }>
  | Readonly<{ status: "loading"; promise: Promise<SettledHyphenationDictionary> }>
  | SettledHyphenationDictionary;

const createInitialStates = (): Record<HyphenationDictionaryId, HyphenationDictionaryState> => ({
  cs: { status: "unloaded" },
  "en-gb": { status: "unloaded" },
  "en-us": { status: "unloaded" },
  sk: { status: "unloaded" },
});

let dictionaryStates = createInitialStates();
let dictionaryGeneration = 0;

// The dictionaries the synchronous layout run in progress asked for and did not
// have. Layout never interleaves, so one slot (saved and restored around nested
// runs) scopes each request to the run, and so to the editor, that made it.
let activeRequests: Set<HyphenationDictionaryId> | undefined;

// Counts every hyphenation that went without its dictionary. A measurement
// taken across a change in this count is incomplete and must not be cached:
// another editor reusing it would never learn that it lacked the dictionary.
let dictionaryMisses = 0;

/** Compare before and after a measurement to tell whether it lacked a dictionary. */
export const getHyphenationDictionaryMisses = (): number => dictionaryMisses;

/** The dictionary that hyphenates text in `locale`, if one is bundled. */
export const hyphenationDictionaryFor = (
  locale: string | undefined,
): HyphenationDictionaryId | undefined => {
  const normalized = locale?.trim().replaceAll("_", "-").toLowerCase();
  if (!normalized) {
    return undefined;
  }
  return HYPHENATION_DICTIONARY_IDS.find(
    (id) => normalized === id || normalized.startsWith(`${id}-`),
  );
};

/** Changes whenever a dictionary finishes loading, so measurement caches cannot go stale. */
export const getHyphenationDictionaryGeneration = (): number => dictionaryGeneration;

const importDictionary = (
  dictionary: HyphenationDictionaryId,
): Promise<SettledHyphenationDictionary> =>
  dictionaryLoaders[dictionary]().then(
    (hyphenate): SettledHyphenationDictionary => ({ status: "loaded", hyphenate }),
    (cause: unknown): SettledHyphenationDictionary => ({
      status: "failed",
      error: new HyphenationDictionaryError({
        message: `The ${dictionary} hyphenation dictionary could not be loaded.`,
        dictionary,
        cause,
      }),
    }),
  );

const startLoad = (dictionary: HyphenationDictionaryId): Promise<SettledHyphenationDictionary> => {
  const promise = importDictionary(dictionary).then((settled) => {
    // A reset (tests) while the import was in flight owns the state now.
    if (dictionaryStates[dictionary].status !== "loading") {
      return settled;
    }
    dictionaryStates[dictionary] = settled;
    switch (settled.status) {
      case "failed":
        recordHyphenationDictionaryError(dictionary, settled.error);
        return settled;
      case "loaded":
        dictionaryGeneration += 1;
        return settled;
      default:
        settled satisfies never;
        return settled;
    }
  });
  dictionaryStates[dictionary] = { status: "loading", promise };
  return promise;
};

/** Load `dictionary` once; later calls share the load or its settled outcome. */
export const requestHyphenationDictionary = (
  dictionary: HyphenationDictionaryId,
): Promise<SettledHyphenationDictionary> => {
  const state = dictionaryStates[dictionary];
  switch (state.status) {
    case "unloaded":
      return startLoad(dictionary);
    case "loading":
      return state.promise;
    case "loaded":
    case "failed":
      return Promise.resolve(state);
    default:
      state satisfies never;
      return state;
  }
};

/**
 * The loaded hyphenator for `dictionary`, or `undefined` while it is not
 * available. An unloaded dictionary starts loading, and a missing one is
 * recorded against the enclosing {@link collectRequestedHyphenationDictionaries}
 * run; the caller hyphenates nothing now.
 */
export const hyphenatorOrRequest = (
  dictionary: HyphenationDictionaryId,
): HyphenateWord | undefined => {
  const state = dictionaryStates[dictionary];
  if (state.status === "loaded") {
    return state.hyphenate;
  }
  dictionaryMisses += 1;
  activeRequests?.add(dictionary);
  if (state.status === "unloaded") {
    void startLoad(dictionary);
  }
  return undefined;
};

export type CollectedHyphenationDictionaries<T> = {
  result: T;
  /** Dictionaries the run needed and did not have: loading or failed. */
  missing: ReadonlySet<HyphenationDictionaryId>;
};

/** Run a synchronous layout pass and report the dictionaries it lacked. */
export const collectRequestedHyphenationDictionaries = <T>(
  run: () => T,
): CollectedHyphenationDictionaries<T> => {
  const outer = activeRequests;
  const missing = new Set<HyphenationDictionaryId>();
  activeRequests = missing;
  try {
    return { result: run(), missing };
  } finally {
    activeRequests = outer;
  }
};

/**
 * Test seam: forget every dictionary so a test observes the unloaded path, and
 * optionally load through `loaders` (for example, ones that fail) until the
 * next reset.
 */
export const resetHyphenationDictionaries = (
  loaders: HyphenationDictionaryLoaders = HYPHENATION_DICTIONARY_LOADERS,
): void => {
  dictionaryLoaders = loaders;
  dictionaryStates = createInitialStates();
  dictionaryGeneration += 1;
};

/** Test seam: which dictionaries have been requested or loaded. */
export const hyphenationDictionaryStatus = (
  dictionary: HyphenationDictionaryId,
): HyphenationDictionaryState["status"] => dictionaryStates[dictionary].status;
