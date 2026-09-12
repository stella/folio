import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import { alignFolioContentBlocks, type FolioContentAlignedBlockEvent } from "./content-alignment";
import type { FolioContentBlock } from "./content-types";

const positionalBlocks = (
  texts: readonly string[],
  side: "base" | "revised",
): FolioContentBlock[] =>
  texts.map((text, index) => ({
    id: `${side}-${String(index)}`,
    idStability: "positional",
    kind: "paragraph",
    text,
  }));

const stableBlock = (id: string, text: string): FolioContentBlock => ({
  id,
  kind: "paragraph",
  text,
});

const reconstruct = (
  events: readonly FolioContentAlignedBlockEvent[],
  side: "base" | "revised",
): FolioContentBlock[] =>
  events.flatMap((event) => {
    switch (event.type) {
      case "pair":
        return [side === "base" ? event.baseBlock : event.revisedBlock];
      case "baseOnly":
        return side === "base" ? [event.block] : [];
      case "revisedOnly":
        return side === "revised" ? [event.block] : [];
      default: {
        const unreachable: never = event;
        return unreachable;
      }
    }
  });

const historicalExactLcs = (base: readonly string[], revised: readonly string[]): string[] => {
  const stride = revised.length + 1;
  const lengths = new Uint16Array((base.length + 1) * stride);
  for (let baseIndex = base.length - 1; baseIndex >= 0; baseIndex--) {
    for (let revisedIndex = revised.length - 1; revisedIndex >= 0; revisedIndex--) {
      lengths[baseIndex * stride + revisedIndex] =
        base[baseIndex] === revised[revisedIndex]
          ? (lengths[(baseIndex + 1) * stride + revisedIndex + 1] ?? 0) + 1
          : Math.max(
              lengths[(baseIndex + 1) * stride + revisedIndex] ?? 0,
              lengths[baseIndex * stride + revisedIndex + 1] ?? 0,
            );
    }
  }

  const selected: string[] = [];
  let baseIndex = 0;
  let revisedIndex = 0;
  while (baseIndex < base.length && revisedIndex < revised.length) {
    const baseText = base[baseIndex];
    if (baseText === revised[revisedIndex] && baseText !== undefined) {
      selected.push(baseText);
      baseIndex += 1;
      revisedIndex += 1;
      continue;
    }
    const down = lengths[(baseIndex + 1) * stride + revisedIndex] ?? 0;
    const right = lengths[baseIndex * stride + revisedIndex + 1] ?? 0;
    if (down >= right) {
      baseIndex += 1;
    } else {
      revisedIndex += 1;
    }
  }
  return selected;
};

describe("gap-local exact block anchors", () => {
  test("does not use whitespace-only text as shifted identity evidence", () => {
    for (const whitespace of ["", " ", "\t\n", "\u00a0"]) {
      const base = positionalBlocks(["Alpha", whitespace, "Beta"], "base");
      const revised = positionalBlocks([whitespace, "Alpha", "Beta"], "revised");

      const events = alignFolioContentBlocks(base, revised);
      expect(events.map(({ type }) => type)).toEqual(["revisedOnly", "pair", "baseOnly", "pair"]);
      expect(reconstruct(events, "base")).toEqual(base);
      expect(reconstruct(events, "revised")).toEqual(revised);
    }
  });

  test("does not promote exact text duplicated in its current gap", () => {
    const repeated = "Confidentiality obligations continue.";
    const base = positionalBlocks([repeated, "Alpha", repeated, "Beta"], "base");
    const revised = positionalBlocks([repeated, repeated, "Alpha", repeated, "Beta"], "revised");

    const events = alignFolioContentBlocks(base, revised);
    expect(events.map(({ type }) => type)).toEqual(["pair", "revisedOnly", "pair", "pair", "pair"]);
    expect(reconstruct(events, "base")).toEqual(base);
    expect(reconstruct(events, "revised")).toEqual(revised);
  });

  test("does not let repeated text displace a later unique anchor", () => {
    const base = positionalBlocks(["Alpha", "Repeated", "Beta", "Repeated", "Gamma"], "base");
    const revised = positionalBlocks(["Repeated", "Alpha", "Repeated", "Beta", "Gamma"], "revised");

    const events = alignFolioContentBlocks(base, revised, { stableIdMismatch: "pair" });
    const exactPairTexts = events.flatMap((event) =>
      event.type === "pair" && event.baseBlock.text === event.revisedBlock.text
        ? [event.baseBlock.text]
        : [],
    );
    expect(exactPairTexts).toEqual(["Alpha", "Repeated", "Beta", "Gamma"]);
    expect(reconstruct(events, "base")).toEqual(base);
    expect(reconstruct(events, "revised")).toEqual(revised);
  });

  test("does not let crossing exact anchors displace stable identity", () => {
    const base = [
      stableBlock("stable-a", "Stable A"),
      ...positionalBlocks(["Exact X", "Exact Y"], "base"),
    ];
    const revised = [
      ...positionalBlocks(["Exact X", "Exact Y"], "revised"),
      stableBlock("stable-a", "Stable A revised"),
    ];

    const events = alignFolioContentBlocks(base, revised);
    expect(events.map(({ type }) => type)).toEqual([
      "revisedOnly",
      "revisedOnly",
      "pair",
      "baseOnly",
      "baseOnly",
    ]);
    expect(
      events.flatMap((event) =>
        event.type === "pair" ? [[event.baseBlock.id, event.revisedBlock.id]] : [],
      ),
    ).toEqual([["stable-a", "stable-a"]]);
    expect(reconstruct(events, "base")).toEqual(base);
    expect(reconstruct(events, "revised")).toEqual(revised);
  });

  test("preserves the exact-LCS crossing tie break", () => {
    const base = positionalBlocks(["Alpha", "Beta"], "base");
    const revised = positionalBlocks(["Beta", "Alpha"], "revised");

    const exactPairs = alignFolioContentBlocks(base, revised, {
      stableIdMismatch: "pair",
    }).flatMap((event) =>
      event.type === "pair" && event.baseBlock.text === event.revisedBlock.text
        ? [event.baseBlock.text]
        : [],
    );
    expect(exactPairs).toEqual(["Beta"]);
  });

  test("preserves the historical exact subsequence for generated unique permutations", () => {
    const texts = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta", "Theta"];
    fc.assert(
      fc.property(
        fc.shuffledSubarray(texts, { minLength: texts.length, maxLength: texts.length }),
        (revisedTexts) => {
          const base = positionalBlocks(texts, "base");
          const revised = positionalBlocks(revisedTexts, "revised");
          const events = alignFolioContentBlocks(base, revised, { stableIdMismatch: "pair" });
          const exactPairs = events.filter(
            (event) => event.type === "pair" && event.baseBlock.text === event.revisedBlock.text,
          );

          expect(exactPairs.map(({ baseBlock }) => baseBlock.text)).toEqual(
            historicalExactLcs(texts, revisedTexts),
          );
          expect(reconstruct(events, "base")).toEqual(base);
          expect(reconstruct(events, "revised")).toEqual(revised);
          expect(alignFolioContentBlocks(base, revised, { stableIdMismatch: "pair" })).toEqual(
            events,
          );
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("reconstructs generated repeated, blank, and Unicode sequences", () => {
    const texts = fc.array(
      fc.constantFrom("", " ", "Alpha", "Repeated terms", "§ 14.2", "😀", "e\u0301", "é"),
      { maxLength: 20 },
    );
    fc.assert(
      fc.property(texts, texts, (baseTexts, revisedTexts) => {
        const base = positionalBlocks(baseTexts, "base");
        const revised = positionalBlocks(revisedTexts, "revised");
        const events = alignFolioContentBlocks(base, revised, { stableIdMismatch: "pair" });

        expect(reconstruct(events, "base")).toEqual(base);
        expect(reconstruct(events, "revised")).toEqual(revised);
        expect(alignFolioContentBlocks(base, revised, { stableIdMismatch: "pair" })).toEqual(
          events,
        );
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });
});
