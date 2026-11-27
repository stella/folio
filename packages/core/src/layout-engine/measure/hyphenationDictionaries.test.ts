import { afterAll, beforeEach, describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import { createRequire } from "node:module";

import type { ParagraphBlock } from "../types";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "./__tests__/fakeTextMeasure";
import { hashParagraphBlock } from "./cache";
import {
  collectRequestedHyphenationDictionaries,
  configureHyphenationAssets,
  HYPHENATION_DICTIONARY_IDS,
  hyphenationDictionaryFor,
  type HyphenationDictionaryId,
  hyphenationDictionaryStatus,
  resetHyphenationDictionaries,
} from "./hyphenationDictionaries";
import { preloadHyphenationDictionaries } from "./hyphenationPreload";
import { findHyphenationBreaks } from "./lineBreaks";
import { measureBlocks } from "./measureBlocks";
import { measureParagraph } from "./measureParagraph";

const DICTIONARY_LOCALES = {
  cs: "cs-CZ",
  "en-gb": "en-GB",
  "en-us": "en-US",
  sk: "sk-SK",
} as const satisfies Record<HyphenationDictionaryId, string>;

const DICTIONARIES = Object.keys(DICTIONARY_LOCALES).filter(
  (id): id is HyphenationDictionaryId => hyphenationDictionaryFor(id) === id,
);

const WORDS = {
  cs: "nejneobhospodářovávatelnější",
  "en-gb": "hyphenation",
  "en-us": "hyphenation",
  sk: "najneobhospodarovávateľnejší",
} as const satisfies Record<HyphenationDictionaryId, string>;

const paragraphs = (attrs?: ParagraphBlock["attrs"]): ParagraphBlock[] =>
  DICTIONARIES.map((id) => ({
    kind: "paragraph",
    id: `lazy-hyphenation-${id}`,
    runs: [
      {
        kind: "text",
        text: WORDS[id],
        language: { val: DICTIONARY_LOCALES[id] },
      },
    ],
    attrs,
  }));

const measureAtNarrowWidth = (blocks: readonly ParagraphBlock[]) => {
  const lines: ReturnType<typeof measureParagraph>["lines"] = [];
  withFakeTextMeasure(
    () => {
      for (const block of blocks) {
        lines.push(...measureParagraph(block, 70).lines);
      }
    },
    { charWidth: fixedCharWidth(10) },
  );
  return lines;
};

const statuses = () =>
  DICTIONARIES.map((id) => hyphenationDictionaryStatus(id));

beforeEach(() => {
  resetHyphenationDictionaries();
});

afterAll(() => {
  resetHyphenationDictionaries();
});

describe("hyphenation dictionaries", () => {
  test("every dictionary is addressed by its own tag and any region or variant of it", () => {
    expect(DICTIONARIES).toHaveLength(Object.keys(DICTIONARY_LOCALES).length);
    for (const id of DICTIONARIES) {
      expect(hyphenationDictionaryFor(DICTIONARY_LOCALES[id])).toBe(id);
      expect(
        hyphenationDictionaryFor(DICTIONARY_LOCALES[id].replace("-", "_")),
      ).toBe(id);
      expect(hyphenationDictionaryFor(` ${id.toUpperCase()}-x-private `)).toBe(
        id,
      );
    }
    expect(hyphenationDictionaryFor("en")).toBeUndefined();
    expect(hyphenationDictionaryFor("csb")).toBeUndefined();
    expect(hyphenationDictionaryFor(undefined)).toBeUndefined();
  });

  test("a paragraph without automatic hyphenation never requests a dictionary", () => {
    const lines = measureAtNarrowWidth(paragraphs());

    expect(lines.length).toBeGreaterThan(1);
    expect(lines.some((line) => line.discretionaryHyphen !== undefined)).toBe(
      false,
    );
    expect(statuses().every((status) => status === "unloaded")).toBe(true);
  });

  test("capitals excluded by doNotHyphenateCaps never request a dictionary", () => {
    expect(
      findHyphenationBreaks("HYPHENATION", {
        locale: "en-US",
        doNotHyphenateCaps: true,
      }),
    ).toEqual([]);
    expect(hyphenationDictionaryStatus("en-us")).toBe("unloaded");
  });

  test("the first request hyphenates nothing, and the load invalidates measurements", async () => {
    const blocks = paragraphs({ automaticHyphenation: { enabled: true } });
    const hashesBeforeLoad = blocks.map(hashParagraphBlock);

    const { result: unloadedLines, missing } =
      collectRequestedHyphenationDictionaries(() =>
        measureAtNarrowWidth(blocks),
      );

    expect(
      unloadedLines.some((line) => line.discretionaryHyphen !== undefined),
    ).toBe(false);
    expect([...missing].toSorted()).toEqual([...DICTIONARIES].toSorted());
    expect(statuses().every((status) => status === "loading")).toBe(true);

    const preloaded = await preloadHyphenationDictionaries(
      Object.values(DICTIONARY_LOCALES),
    );

    expect(preloaded.isOk()).toBe(true);
    expect(statuses().every((status) => status === "loaded")).toBe(true);
    expect(
      collectRequestedHyphenationDictionaries(() =>
        measureAtNarrowWidth(blocks),
      ).missing.size,
    ).toBe(0);
    for (const [index, block] of blocks.entries()) {
      expect(hashParagraphBlock(block)).not.toBe(hashesBeforeLoad[index]);
      expect(measureAtNarrowWidth([block]).at(0)?.discretionaryHyphen).toEqual({
        runIndex: 0,
      });
    }
  });

  test("a measurement that lacked a dictionary is not cached for the next run", async () => {
    const blocks = paragraphs({ automaticHyphenation: { enabled: true } });
    const missingPerRun: number[] = [];
    withFakeTextMeasure(
      () => {
        for (let run = 0; run < 2; run += 1) {
          missingPerRun.push(
            collectRequestedHyphenationDictionaries(() =>
              measureBlocks(blocks, 70),
            ).missing.size,
          );
        }
      },
      { charWidth: fixedCharWidth(10) },
    );

    expect(missingPerRun).toEqual([DICTIONARIES.length, DICTIONARIES.length]);
    expect(
      (
        await preloadHyphenationDictionaries(Object.values(DICTIONARY_LOCALES))
      ).isOk(),
    ).toBe(true);
  });

  test("a loaded dictionary hyphenates exactly as a preloaded one", async () => {
    const preloaded = await preloadHyphenationDictionaries(["cs-CZ"]);

    expect(preloaded.isOk()).toBe(true);
    expect(
      findHyphenationBreaks("nejneobhospodářovávatelnější", {
        locale: "cs-CZ",
      }),
    ).toEqual([3, 5, 7, 10, 12, 14, 16, 18, 20, 23, 26]);
    expect(hyphenationDictionaryStatus("en-us")).toBe("unloaded");
  });
});

describe("hyphenation dictionary assets", () => {
  const require = createRequire(import.meta.url);
  const nowhere = new URL("file:///nonexistent-folio-hyphenation/");
  const packaged = (id: HyphenationDictionaryId): Promise<unknown> =>
    readFile(
      new URL(`../../../assets/hyphenation/${id}.json`, import.meta.url),
      "utf8",
    ).then((json): unknown => JSON.parse(json));

  test("the packaged files are the pinned patterns", async () => {
    for (const id of HYPHENATION_DICTIONARY_IDS) {
      expect(await packaged(id)).toEqual(
        require(`hyphen/patterns/${id}.js`) as unknown,
      );
    }
  });

  test("a dictionary handed over is used without reading anything", async () => {
    configureHyphenationAssets({
      baseUrl: nowhere,
      dictionaries: { cs: await packaged("cs") },
    });

    expect((await preloadHyphenationDictionaries(["cs-CZ"])).isOk()).toBe(true);
    expect(
      findHyphenationBreaks("nejneobhospodářovávatelnější", {
        locale: "cs-CZ",
      }),
    ).toEqual([3, 5, 7, 10, 12, 14, 16, 18, 20, 23, 26]);
  });

  test("an unreadable or malformed dictionary fails, and new options retry it", async () => {
    configureHyphenationAssets({ baseUrl: nowhere, dictionaries: { sk: 42 } });

    const unreadable = await preloadHyphenationDictionaries(["en-US"]);
    const malformed = await preloadHyphenationDictionaries(["sk-SK"]);
    expect(unreadable.isErr() && unreadable.error.dictionary).toBe("en-us");
    expect(malformed.isErr() && malformed.error.dictionary).toBe("sk");
    expect(findHyphenationBreaks("hyphenation", { locale: "en-US" })).toEqual(
      [],
    );

    configureHyphenationAssets({});
    expect(hyphenationDictionaryStatus("en-us")).toBe("unloaded");
    expect((await preloadHyphenationDictionaries(["en-US"])).isOk()).toBe(true);
  });
});
