import { afterAll, expect, test } from "bun:test";

import type { LayoutInstrumentation } from "../layoutInstrumentation";
import {
  type HyphenationDictionaryEvent,
  type HyphenationDictionaryLoaders,
  hyphenationDictionaryStatus,
  onHyphenationDictionarySettled,
  resetHyphenationDictionaries,
} from "./hyphenationDictionaries";
import { preloadHyphenationDictionaries } from "./hyphenationPreload";
import { findHyphenationBreaks } from "./lineBreaks";

// A chunk that cannot be fetched, as after a deployment replaced it.
let attempts = 0;
const missingChunk = async (): Promise<never> => {
  attempts += 1;
  throw new Error("chunk missing");
};
const MISSING_CHUNKS = {
  cs: missingChunk,
  "en-gb": missingChunk,
  "en-us": missingChunk,
  sk: missingChunk,
} as const satisfies HyphenationDictionaryLoaders;

type RecordedError = Parameters<
  NonNullable<LayoutInstrumentation["onHyphenationDictionaryError"]>
>[0];

const previousInstrumentation = globalThis.__folioLayoutInstrumentation;

afterAll(() => {
  globalThis.__folioLayoutInstrumentation = previousInstrumentation;
  resetHyphenationDictionaries();
});

test("a failed load is reported to subscribers and telemetry, and never retried", async () => {
  resetHyphenationDictionaries(MISSING_CHUNKS);
  const events: HyphenationDictionaryEvent[] = [];
  const recorded: RecordedError[] = [];
  globalThis.__folioLayoutInstrumentation = {
    onHyphenationDictionaryError: (event) => recorded.push(event),
  };
  const unsubscribe = onHyphenationDictionarySettled((event) => events.push(event));

  expect(findHyphenationBreaks("najneobhospodarovávateľnejší", { locale: "sk-SK" })).toEqual([]);
  const preloaded = await preloadHyphenationDictionaries(["sk-SK"]);
  expect(findHyphenationBreaks("najneobhospodarovávateľnejší", { locale: "sk-SK" })).toEqual([]);
  const preloadedAgain = await preloadHyphenationDictionaries(["sk-SK"]);
  unsubscribe();

  expect(preloaded.isErr()).toBe(true);
  expect(preloadedAgain.isErr()).toBe(true);
  expect(attempts).toBe(1);
  expect(events).toHaveLength(1);
  expect(events.at(0)).toMatchObject({ type: "failed", dictionary: "sk" });
  expect(recorded).toEqual([
    { dictionary: "sk", message: "The sk hyphenation dictionary could not be loaded." },
  ]);
  expect(hyphenationDictionaryStatus("sk")).toBe("failed");
});
