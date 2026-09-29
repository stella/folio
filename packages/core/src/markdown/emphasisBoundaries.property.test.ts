/**
 * Emphasis at any boundary survives the markdown writer and reader. Runs of
 * words, punctuation and spaces, bold, italic or struck in any combination,
 * write markdown that reads back as the same text, with no emphasis the runs
 * did not have: the writer never writes a delimiter the reader takes for a
 * literal asterisk. Emphasis on whole words, or on either part of a word,
 * reads back on every letter; the writer may leave punctuation or a space at
 * an emphasis edge outside it where no delimiter could open or close there.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { Document, Paragraph, Run, TextFormatting } from "../types/document";
import { fromMarkdown } from "./fromMarkdown";
import { toMarkdown } from "./index";

setDefaultTimeout(propertyTestTimeout(30_000));

const CLEAN = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  hyperlinks: "inline",
  footnotes: "strip",
} as const;

type Segment = { text: string; bold: boolean; italic: boolean; strike: boolean };

const run = ({ text, bold, italic, strike }: Segment): Run => {
  const formatting: TextFormatting = {
    ...(bold && { bold: true }),
    ...(italic && { italic: true }),
    ...(strike && { strike: true }),
  };
  return {
    type: "run",
    content: [{ type: "text", text }],
    ...(Object.keys(formatting).length > 0 && { formatting }),
  };
};

/** A paragraph framed by plain words, so no segment starts a block construct. */
const documentOf = (segments: readonly Segment[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            run({ text: "Start ", bold: false, italic: false, strike: false }),
            ...segments.map(run),
            run({ text: " end", bold: false, italic: false, strike: false }),
          ],
        },
      ],
    },
  },
});

type Character = { char: string; bold: boolean; italic: boolean; strike: boolean };

const charactersOf = (paragraph: Paragraph): Character[] =>
  paragraph.content.flatMap((item) =>
    item.type === "run"
      ? item.content.flatMap((content) =>
          content.type === "text"
            ? Array.from(content.text, (char) => ({
                char,
                bold: item.formatting?.bold === true,
                italic: item.formatting?.italic === true,
                strike: item.formatting?.strike === true,
              }))
            : [],
        )
      : [],
  );

const segmentOf = (chars: readonly string[]): fc.Arbitrary<Segment> =>
  fc.record({
    text: fc
      .array(fc.constantFrom(...chars), { minLength: 1, maxLength: 4 })
      .map((picked) => picked.join("")),
    bold: fc.boolean(),
    italic: fc.boolean(),
    strike: fc.boolean(),
  });

const WORDS = ["a", "b", "Z", "7", " ", "é"];
const WORDS_AND_PUNCTUATION = [...WORDS, ".", ",", "!", "(", ")", "'", "*", "_", "~"];

const isLetter = (char: string) => /[\p{L}\p{N}]/u.test(char);

const roundTrip = (segments: readonly Segment[]) => {
  const source = documentOf(segments);
  const markdown = toMarkdown(source, CLEAN);
  const [paragraph] = fromMarkdown(markdown).package.document.content;
  if (paragraph?.type !== "paragraph") throw new Error(`no paragraph read from ${markdown}`);
  const [written] = source.package.document.content;
  if (written?.type !== "paragraph") throw new Error("no paragraph written");
  return { markdown, before: charactersOf(written), after: charactersOf(paragraph) };
};

const textOf = (characters: readonly Character[]) => characters.map(({ char }) => char).join("");

describe("markdown emphasis boundaries (properties)", () => {
  test("emphasis on words and on either part of a word reads back as written", () => {
    // Words apart, each emphasized whole or in two parts. (Three differently
    // emphasized parts of one word can overlap in a way the reader cannot
    // match; the next property covers any mix for its text.)
    const word = fc.array(segmentOf(["a", "b", "Z", "7", "é"]), { minLength: 1, maxLength: 2 });
    const plainSpace: Segment = { text: " ", bold: false, italic: false, strike: false };
    const words = fc
      .array(word, { minLength: 1, maxLength: 4 })
      .map((list) =>
        list.flatMap((parts, index) => (index === 0 ? parts : [plainSpace, ...parts])),
      );
    fc.assert(
      fc.property(words, (segments) => {
        const { markdown, before, after } = roundTrip(segments);
        expect({ markdown, text: textOf(after) }).toEqual({ markdown, text: textOf(before) });
        for (const [index, written] of before.entries()) {
          const read = after[index];
          if (!read || !isLetter(written.char)) continue;
          expect({ markdown, index, ...read }).toEqual({ markdown, index, ...written });
        }
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("beside punctuation, the text reads back and no emphasis is added", () => {
    fc.assert(
      fc.property(
        fc.array(segmentOf(WORDS_AND_PUNCTUATION), { minLength: 1, maxLength: 6 }),
        (segments) => {
          const { markdown, before, after } = roundTrip(segments);
          expect({ markdown, text: textOf(after) }).toEqual({ markdown, text: textOf(before) });
          for (const [index, written] of before.entries()) {
            const read = after[index];
            if (!read) continue;
            for (const key of ["bold", "italic", "strike"] as const) {
              if (read[key])
                expect({ markdown, index, key, written: written[key] }).toEqual({
                  markdown,
                  index,
                  key,
                  written: true,
                });
            }
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("emphasis ending in punctuation right before a word reads back without a stray delimiter", () => {
    const { markdown, after } = roundTrip([
      { text: "Closing.", bold: false, italic: true, strike: false },
      { text: "revised", bold: false, italic: false, strike: false },
    ]);
    expect(after.map(({ char }) => char).join("")).toBe("Start Closing.revised end");
    expect(after.find(({ char }) => char === "C")?.italic).toBe(true);
    expect(markdown).not.toContain("\\*");
  });
});
