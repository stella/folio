import { afterAll, beforeEach, describe, expect, test } from "bun:test";

import type { LayoutInstrumentation } from "../layout-engine/layoutInstrumentation";
import {
  collectRequestedHyphenationDictionaries,
  type HyphenateWord,
  type HyphenationDictionaryLoaders,
  hyphenationDictionaryStatus,
  requestHyphenationDictionary,
  resetHyphenationDictionaries,
} from "../layout-engine/measure/hyphenationDictionaries";
import { findHyphenationBreaks } from "../layout-engine/measure/lineBreaks";
import { createHyphenationReadiness } from "./hyphenationReadiness";

const WORD = "najneobhospodarovávateľnejší";

let attempts = 0;
// A chunk that cannot be fetched, as after a deployment replaced it.
const missingChunk = async (): Promise<HyphenateWord> => {
  attempts += 1;
  throw new Error("chunk missing");
};
const identity: HyphenateWord = (text) => text;
const available = async (): Promise<HyphenateWord> => {
  attempts += 1;
  return identity;
};

const loadersWith = (load: () => Promise<HyphenateWord>) =>
  ({
    cs: load,
    "en-gb": load,
    "en-us": load,
    sk: load,
  }) as const satisfies HyphenationDictionaryLoaders;

type RecordedError = Parameters<
  NonNullable<LayoutInstrumentation["onHyphenationDictionaryError"]>
>[0];

/** One editor: counts its relayouts and the errors reported to it. */
const editor = () => {
  const seen = { relayouts: 0, errors: [] as string[] };
  const readiness = createHyphenationReadiness({
    relayout: () => {
      seen.relayouts += 1;
    },
    onError: (error) => seen.errors.push(error.dictionary),
  });
  /** A layout run that hyphenates `locale`, handing the pipeline's result to this editor. */
  const layOut = (locale: string) => {
    const { missing } = collectRequestedHyphenationDictionaries(() =>
      findHyphenationBreaks(WORD, { locale }),
    );
    readiness.track(missing);
  };
  return { seen, readiness, layOut };
};

const settle = async (): Promise<void> => {
  await requestHyphenationDictionary("sk");
  await Promise.resolve();
};

const previousInstrumentation = globalThis.__folioLayoutInstrumentation;

beforeEach(() => {
  attempts = 0;
});

afterAll(() => {
  globalThis.__folioLayoutInstrumentation = previousInstrumentation;
  resetHyphenationDictionaries();
});

describe("createHyphenationReadiness", () => {
  test("re-lays out only the editor whose run lacked the dictionary", async () => {
    resetHyphenationDictionaries(loadersWith(available));
    const requester = editor();
    const bystander = editor();

    requester.layOut("sk-SK");
    bystander.layOut("fr-FR");
    await settle();

    expect(requester.seen).toEqual({ relayouts: 1, errors: [] });
    expect(bystander.seen).toEqual({ relayouts: 0, errors: [] });
    expect(hyphenationDictionaryStatus("sk")).toBe("loaded");

    // The re-run finds the dictionary loaded, so nothing is left to follow.
    requester.layOut("sk-SK");
    await settle();
    expect(requester.seen.relayouts).toBe(1);
    expect(attempts).toBe(1);
  });

  test("reports a failed load once per editor that needs it, including later ones", async () => {
    resetHyphenationDictionaries(loadersWith(missingChunk));
    const recorded: RecordedError[] = [];
    globalThis.__folioLayoutInstrumentation = {
      onHyphenationDictionaryError: (event) => recorded.push(event),
    };
    const requester = editor();
    const bystander = editor();

    requester.layOut("sk-SK");
    bystander.layOut("fr-FR");
    await settle();
    requester.layOut("sk-SK");
    await settle();
    const latecomer = editor();
    latecomer.layOut("sk-SK");
    await settle();

    expect(requester.seen).toEqual({ relayouts: 0, errors: ["sk"] });
    expect(bystander.seen).toEqual({ relayouts: 0, errors: [] });
    expect(latecomer.seen).toEqual({ relayouts: 0, errors: ["sk"] });
    expect(attempts).toBe(1);
    expect(recorded).toEqual([
      { dictionary: "sk", message: "The sk hyphenation dictionary could not be loaded." },
    ]);
    expect(hyphenationDictionaryStatus("sk")).toBe("failed");
  });

  test("a cancelled follow-up is ignored, and a later run is followed afresh", async () => {
    resetHyphenationDictionaries(loadersWith(available));
    const unmounted = editor();
    const remounted = editor();

    unmounted.layOut("sk-SK");
    unmounted.readiness.cancel();
    remounted.layOut("sk-SK");
    remounted.readiness.cancel();
    remounted.layOut("sk-SK");
    await settle();

    expect(unmounted.seen).toEqual({ relayouts: 0, errors: [] });
    expect(remounted.seen).toEqual({ relayouts: 1, errors: [] });
  });
});
