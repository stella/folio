/**
 * On-demand hyphenation dictionaries for automatic hyphenation.
 *
 * Each `hyphen` pattern set is tens to hundreds of kilobytes of source whose
 * evaluation builds a pattern trie, and most documents never enable
 * `w:autoHyphenation`. A dictionary is therefore imported only once measurement
 * asks to hyphenate a word in its language. Measurement is synchronous, so the
 * first request starts the import and hyphenates nothing; subscribers to
 * {@link onHyphenationDictionarySettled} re-run layout once it loads and
 * surface the error if it fails. Hosts
 * that must lay out correctly on the first pass (headless layout, tests) await
 * `preloadHyphenationDictionaries` (`./hyphenationPreload`) before measuring.
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

export type HyphenationDictionaryEvent =
  | Readonly<{ type: "loaded"; dictionary: HyphenationDictionaryId }>
  | Readonly<{
      type: "failed";
      dictionary: HyphenationDictionaryId;
      error: HyphenationDictionaryError;
    }>;

type HyphenationDictionaryListener = (event: HyphenationDictionaryEvent) => void;

const listeners = new Set<HyphenationDictionaryListener>();

const notify = (event: HyphenationDictionaryEvent): void => {
  for (const listener of listeners) {
    listener(event);
  }
};

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
        notify({ type: "failed", dictionary, error: settled.error });
        return settled;
      case "loaded":
        dictionaryGeneration += 1;
        notify({ type: "loaded", dictionary });
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
 * available. An unloaded dictionary starts loading; the caller hyphenates
 * nothing now and is re-run through {@link onHyphenationDictionarySettled}.
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
 * Subscribe to settled dictionary loads; returns the unsubscribe function. A
 * failure is terminal for the session, so a host surfaces it rather than
 * waiting for a relayout that will not come.
 */
export const onHyphenationDictionarySettled = (
  listener: HyphenationDictionaryListener,
): (() => void) => {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
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
