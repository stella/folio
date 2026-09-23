/**
 * On-demand hyphenation dictionaries for automatic hyphenation.
 *
 * Each `hyphen` pattern set is tens to hundreds of kilobytes of source whose
 * evaluation builds a pattern trie, and most documents never enable
 * `w:autoHyphenation`. A dictionary is therefore imported only once measurement
 * asks to hyphenate a word in its language. Measurement is synchronous, so the
 * first request starts the import and hyphenates nothing; subscribers to
 * {@link onHyphenationDictionaryLoaded} re-run layout once it resolves. Hosts
 * that must lay out correctly on the first pass (headless layout, tests) await
 * {@link preloadHyphenationDictionaries} before measuring.
 */

import { Result, TaggedError } from "better-result";

import { recordHyphenationDictionaryError } from "../layoutInstrumentation";

export type HyphenateWord = (text: string) => string;

const HYPHENATION_DICTIONARY_IDS = ["cs", "en-gb", "en-us", "sk"] as const;

export type HyphenationDictionaryId = (typeof HYPHENATION_DICTIONARY_IDS)[number];

// Literal specifiers, one per dictionary, so bundlers emit one lazy chunk each.
const HYPHENATION_DICTIONARY_LOADERS = {
  cs: async () => (await import("hyphen/cs")).default.hyphenateSync,
  "en-gb": async () => (await import("hyphen/en-gb")).default.hyphenateSync,
  "en-us": async () => (await import("hyphen/en-us")).default.hyphenateSync,
  sk: async () => (await import("hyphen/sk")).default.hyphenateSync,
} as const satisfies Record<HyphenationDictionaryId, () => Promise<HyphenateWord>>;

export class HyphenationDictionaryError extends TaggedError("HyphenationDictionaryError")<{
  message: string;
  dictionary: HyphenationDictionaryId;
  cause: unknown;
}> {}

type HyphenationDictionaryLoad = Result<HyphenateWord, HyphenationDictionaryError>;

type HyphenationDictionaryState =
  | Readonly<{ status: "unloaded" }>
  | Readonly<{ status: "loading"; promise: Promise<HyphenationDictionaryLoad> }>
  | Readonly<{ status: "loaded"; hyphenate: HyphenateWord }>
  // Terminal for the session: retrying on every layout would turn a missing
  // chunk into an endless request/relayout loop.
  | Readonly<{ status: "failed"; error: HyphenationDictionaryError }>;

const createInitialStates = (): Record<HyphenationDictionaryId, HyphenationDictionaryState> => ({
  cs: { status: "unloaded" },
  "en-gb": { status: "unloaded" },
  "en-us": { status: "unloaded" },
  sk: { status: "unloaded" },
});

let dictionaryStates = createInitialStates();
let dictionaryGeneration = 0;

type HyphenationDictionaryLoadedListener = (dictionary: HyphenationDictionaryId) => void;

const loadedListeners = new Set<HyphenationDictionaryLoadedListener>();

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

const startLoad = (dictionary: HyphenationDictionaryId): Promise<HyphenationDictionaryLoad> => {
  const promise = Result.tryPromise({
    try: HYPHENATION_DICTIONARY_LOADERS[dictionary],
    catch: (cause) =>
      new HyphenationDictionaryError({
        message: `The ${dictionary} hyphenation dictionary could not be loaded.`,
        dictionary,
        cause,
      }),
  }).then((loaded) => {
    // A reset (tests) while the import was in flight owns the state now.
    if (dictionaryStates[dictionary].status !== "loading") {
      return loaded;
    }
    if (loaded.isErr()) {
      dictionaryStates[dictionary] = { status: "failed", error: loaded.error };
      recordHyphenationDictionaryError(dictionary, loaded.error);
      return loaded;
    }
    dictionaryStates[dictionary] = { status: "loaded", hyphenate: loaded.value };
    dictionaryGeneration += 1;
    for (const listener of loadedListeners) {
      listener(dictionary);
    }
    return loaded;
  });
  dictionaryStates[dictionary] = { status: "loading", promise };
  return promise;
};

const requestDictionary = (
  dictionary: HyphenationDictionaryId,
): Promise<HyphenationDictionaryLoad> => {
  const state = dictionaryStates[dictionary];
  switch (state.status) {
    case "unloaded":
      return startLoad(dictionary);
    case "loading":
      return state.promise;
    case "loaded":
      return Promise.resolve(Result.ok(state.hyphenate));
    case "failed":
      return Promise.resolve(Result.err(state.error));
    default:
      state satisfies never;
      return state;
  }
};

/**
 * The loaded hyphenator for `dictionary`, or `undefined` while it is not
 * available. An unloaded dictionary starts loading; the caller hyphenates
 * nothing now and is re-run through {@link onHyphenationDictionaryLoaded}.
 */
export const hyphenatorOrRequest = (
  dictionary: HyphenationDictionaryId,
): HyphenateWord | undefined => {
  const state = dictionaryStates[dictionary];
  switch (state.status) {
    case "loaded":
      return state.hyphenate;
    case "unloaded":
      void startLoad(dictionary);
      return undefined;
    case "loading":
    case "failed":
      return undefined;
    default:
      state satisfies never;
      return state;
  }
};

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
  const loads = await Promise.all([...dictionaries].map(requestDictionary));
  const failed = loads.find((load) => load.isErr());
  return failed?.isErr() ? Result.err(failed.error) : Result.ok(undefined);
};

/** Subscribe to dictionary loads; returns the unsubscribe function. */
export const onHyphenationDictionaryLoaded = (
  listener: HyphenationDictionaryLoadedListener,
): (() => void) => {
  loadedListeners.add(listener);
  return () => {
    loadedListeners.delete(listener);
  };
};

/** Test seam: forget every dictionary so a test observes the unloaded path. */
export const resetHyphenationDictionaries = (): void => {
  dictionaryStates = createInitialStates();
  dictionaryGeneration += 1;
};

/** Test seam: which dictionaries have been requested or loaded. */
export const hyphenationDictionaryStatus = (
  dictionary: HyphenationDictionaryId,
): HyphenationDictionaryState["status"] => dictionaryStates[dictionary].status;
