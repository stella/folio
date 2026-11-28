/**
 * On-demand hyphenation dictionaries for automatic hyphenation.
 *
 * Each pattern set is tens of kilobytes, and most documents never enable
 * `w:autoHyphenation`. The sets therefore ship as JSON data beside the package
 * (`assets/hyphenation/<id>.json`, exported as
 * `@stll/folio-core/hyphenation/<id>.json`) rather than as modules, so no
 * bundle carries them, and a dictionary is fetched only once measurement asks
 * to hyphenate a word in its language. A host that cannot fetch them from
 * there (a single-file bundle, offline, a strict `connect-src`) points
 * {@link configureHyphenationAssets} at another location or hands the data
 * over. Measurement is synchronous, so the first request starts the load and
 * hyphenates nothing. A layout run wrapped
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

import { TaggedError, type TaggedErrorClass } from "better-result";

import { readRuntimeJsonAsset } from "../../utils/runtimeJsonAsset";
import { recordHyphenationDictionaryError } from "../layoutInstrumentation";

export type HyphenateWord = (text: string) => string;

export type HyphenationDictionaryId = "cs" | "en-gb" | "en-us" | "sk";

/** The languages folio hyphenates, by dictionary. */
export const HYPHENATION_DICTIONARY_IDS: readonly HyphenationDictionaryId[] = [
  "cs",
  "en-gb",
  "en-us",
  "sk",
];

export type HyphenationDictionaryLoaders = Readonly<
  Record<HyphenationDictionaryId, () => Promise<HyphenateWord>>
>;

/**
 * One dictionary's patterns: the parsed contents of
 * `@stll/folio-core/hyphenation/<id>.json`, passed on unchanged.
 */
export type HyphenationPatterns = unknown;

export type HyphenationAssetOptions = {
  /**
   * The absolute URL of a directory holding the package's
   * `hyphenation/<id>.json` files, to fetch from instead of beside the package
   * (a CDN path of the same version, say).
   */
  readonly baseUrl?: string | URL;
  /**
   * Dictionaries the host already holds, used without fetching anything:
   * offline hosts, or pages whose Content Security Policy blocks the fetch.
   */
  readonly dictionaries?: Readonly<
    Partial<Record<HyphenationDictionaryId, HyphenationPatterns>>
  >;
};

let assetOptions: HyphenationAssetOptions = {};

// One literal `new URL` per dictionary, so bundlers that copy referenced files
// emit each as an asset. The path holds from `src/` and `dist/` alike.
const packagedDictionaryUrl = (dictionary: HyphenationDictionaryId): URL => {
  switch (dictionary) {
    case "cs":
      return new URL("../../../assets/hyphenation/cs.json", import.meta.url);
    case "en-gb":
      return new URL("../../../assets/hyphenation/en-gb.json", import.meta.url);
    case "en-us":
      return new URL("../../../assets/hyphenation/en-us.json", import.meta.url);
    case "sk":
      return new URL("../../../assets/hyphenation/sk.json", import.meta.url);
    default:
      dictionary satisfies never;
      return dictionary;
  }
};

const dictionaryUrl = (dictionary: HyphenationDictionaryId): URL => {
  const { baseUrl } = assetOptions;
  if (baseUrl === undefined) {
    return packagedDictionaryUrl(dictionary);
  }
  const url = new URL(baseUrl);
  if (!url.pathname.endsWith("/")) {
    url.pathname += "/";
  }
  url.pathname += `${dictionary}.json`;
  return url;
};

const hyphenatorFor = async (
  patterns: HyphenationPatterns,
): Promise<HyphenateWord> => {
  const { default: createHyphenator } = await import("hyphen");
  // The data is the `hyphen` pattern module's value as JSON; its declared
  // type does not describe that shape.
  const hyphenate = createHyphenator(
    patterns as Parameters<typeof createHyphenator>[0],
  ) as HyphenateWord;
  // Data that is not a pattern set fails here, as a failed load, rather
  // than inside a layout pass.
  hyphenate("hyphenation");
  return hyphenate;
};

const loadDictionary = async (
  dictionary: HyphenationDictionaryId,
): Promise<HyphenateWord> =>
  hyphenatorFor(
    assetOptions.dictionaries?.[dictionary] ??
      (await readRuntimeJsonAsset(dictionaryUrl(dictionary))),
  );

const HYPHENATION_DICTIONARY_LOADERS = {
  cs: () => loadDictionary("cs"),
  "en-gb": () => loadDictionary("en-gb"),
  "en-us": () => loadDictionary("en-us"),
  sk: () => loadDictionary("sk"),
} as const satisfies HyphenationDictionaryLoaders;

let dictionaryLoaders: HyphenationDictionaryLoaders =
  HYPHENATION_DICTIONARY_LOADERS;

const HyphenationDictionaryErrorBase: TaggedErrorClass<"HyphenationDictionaryError"> =
  TaggedError("HyphenationDictionaryError");

export class HyphenationDictionaryError extends HyphenationDictionaryErrorBase<{
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
  | Readonly<{
      status: "loading";
      promise: Promise<SettledHyphenationDictionary>;
    }>
  | SettledHyphenationDictionary;

const createInitialStates = (): Record<
  HyphenationDictionaryId,
  HyphenationDictionaryState
> => ({
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
export const getHyphenationDictionaryGeneration = (): number =>
  dictionaryGeneration;

const importDictionary = (
  dictionary: HyphenationDictionaryId,
): Promise<SettledHyphenationDictionary> =>
  dictionaryLoaders[dictionary]().then(
    (hyphenate): SettledHyphenationDictionary => ({
      status: "loaded",
      hyphenate,
    }),
    (cause: unknown): SettledHyphenationDictionary => ({
      status: "failed",
      error: new HyphenationDictionaryError({
        message: `The ${dictionary} hyphenation dictionary could not be loaded.`,
        dictionary,
        cause,
      }),
    }),
  );

const startLoad = (
  dictionary: HyphenationDictionaryId,
): Promise<SettledHyphenationDictionary> => {
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
 * Set where hyphenation dictionaries come from; each call replaces the
 * previous options. Call it before opening documents. A dictionary that failed
 * to load is tried again under the new options; loaded ones are kept.
 */
export const configureHyphenationAssets = (
  options: HyphenationAssetOptions,
): void => {
  assetOptions = options;
  for (const dictionary of HYPHENATION_DICTIONARY_IDS) {
    if (dictionaryStates[dictionary].status === "failed") {
      dictionaryStates[dictionary] = { status: "unloaded" };
    }
  }
};

/**
 * Test seam: forget every dictionary and the asset options so a test observes
 * the unloaded path, and optionally load through `loaders` (for example, ones
 * that fail) until the next reset.
 */
export const resetHyphenationDictionaries = (
  loaders: HyphenationDictionaryLoaders = HYPHENATION_DICTIONARY_LOADERS,
): void => {
  dictionaryLoaders = loaders;
  assetOptions = {};
  dictionaryStates = createInitialStates();
  dictionaryGeneration += 1;
};

/** Test seam: which dictionaries have been requested or loaded. */
export const hyphenationDictionaryStatus = (
  dictionary: HyphenationDictionaryId,
): HyphenationDictionaryState["status"] => dictionaryStates[dictionary].status;
