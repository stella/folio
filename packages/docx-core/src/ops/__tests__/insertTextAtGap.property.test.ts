import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  ParagraphContent,
  Run,
  RunContent,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { equalForStaleness } from "../equality";
import { compareGaps, type Gap, type InlineNode, leafSpans } from "../leaves";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertedRunProps,
  type OpStory,
} from "../types";
import { RUN_LEAF_SAMPLES, RUN_LEAVES } from "./runLeafSamples";

const BLOCK = "00000001";

/**
 * The footnote a paragraph holding a note's own mark lies in: not the note
 * the sampled note references name.
 */
const NOTE = 9;

/** The paragraph in the body, or in a footnote when it holds a note's own mark. */
const placeOf = (
  content: readonly ParagraphContent[],
): { story: OpStory; document: Document } => {
  const paragraph = {
    type: "paragraph" as const,
    paraId: BLOCK,
    content: [...content],
  };
  const holdsMark = leafSpans(content).some(
    ({ node }) => node.type === "noteMarker",
  );
  if (!holdsMark)
    return {
      story: "main",
      document: { package: { document: { content: [paragraph] } } },
    };
  return {
    story: { kind: "footnote", id: NOTE },
    document: {
      package: {
        document: {
          content: [{ type: "paragraph", paraId: "0000000F", content: [] }],
        },
        footnotes: [{ type: "footnote", id: NOTE, content: [paragraph] }],
      },
    },
  };
};

const paragraphOf = (document: Document, story: OpStory) => {
  const blocks =
    story === "main"
      ? document.package.document.content
      : (document.package.footnotes?.at(0)?.content ?? []);
  const paragraph = blocks.find((block) => block.type === "paragraph");
  if (paragraph?.type !== "paragraph")
    throw new Error("The sampled paragraph is there.");
  return paragraph.content;
};

/** Every gap of the content, in order. */
const gapsOf = (content: readonly InlineNode[]): Gap[] => {
  const seen = new Set<string>();
  return leafSpans(content)
    .flatMap(({ before, after }) => [before, after])
    .filter((gap) => {
      const key = `${String(gap.offset)}/${String(gap.zeroWidthBefore)}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .toSorted(compareGaps);
};

/** One token per character and per other leaf, with the run holding it. */
type Token = { name: string; run: Run | undefined; after: Gap };

const tokensOf = (content: readonly InlineNode[]): Token[] =>
  leafSpans(content).flatMap(({ node, ancestors, before, after }) => {
    const holder = ancestors.at(-1);
    const run = holder?.type === "run" ? holder : undefined;
    if (node.type === "text" && node.text !== "")
      return [...node.text].map((character, index) => ({
        name: character,
        run,
        after: { offset: before.offset + index + 1, zeroWidthBefore: 0 },
      }));
    return [{ name: `<${node.type}>`, run, after }];
  });

/** Which neighbour the typed text shares a run with. */
const hostSide = (
  withBefore: boolean,
  withAfter: boolean,
  alone: boolean,
): string => {
  if (withBefore) return "with the leaf before";
  if (withAfter) return "with the leaf after";
  return alone ? "alone" : "elsewhere";
};

const TYPED = "XY";
const ITALIC = { italic: true } as const;

/**
 * How typing at a gap must land, or a description of how it did not: every
 * leaf in its place with the text exactly at the gap; inheriting text joins
 * the run holding the leaf before the gap, else the one holding the leaf
 * after it, else a run of its own; formatted text is a run of its own; and
 * the inverse restores the paragraph.
 */
const landing = (
  base: Document,
  story: OpStory,
  gap: Gap,
  runProps: InsertedRunProps,
  applied: { document: Document; inverse: readonly DocumentOp[] } | string,
): object | undefined => {
  if (typeof applied === "string") return { refused: applied };
  const before = tokensOf(paragraphOf(base, story));
  const after = tokensOf(paragraphOf(applied.document, story));
  const at = before.filter(
    (token) => compareGaps(token.after, gap) <= 0,
  ).length;
  const expected = [
    ...before.slice(0, at).map(({ name }) => name),
    ...TYPED,
    ...before.slice(at).map(({ name }) => name),
  ];
  const names = after.map(({ name }) => name);
  if (JSON.stringify(names) !== JSON.stringify(expected))
    return { leaves: names, expected };
  const typed = after.slice(at, at + TYPED.length);
  const host = typed[0]?.run;
  if (host === undefined || typed.some(({ run }) => run !== host))
    return { typedIn: "no one run" };
  const previous = after[at - 1];
  const next = after[at + TYPED.length];
  const own = host.content;
  const alone =
    own.length === 1 && own[0]?.type === "text" && own[0].text === TYPED;
  if (runProps !== "inherit") {
    if (!alone || host.formatting?.italic !== true)
      return { formattedTextIn: host };
  } else {
    const wanted = hostSide(
      before[at - 1]?.run !== undefined,
      before[at]?.run !== undefined,
      true,
    );
    const got = hostSide(previous?.run === host, next?.run === host, alone);
    if (got !== wanted) return { host: got, wanted };
  }
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) return { inverse: restored.error.message };
  return equalForStaleness(restored.value.document, base)
    ? undefined
    : { restored: paragraphOf(restored.value.document, story) };
};

const typeAt = (
  base: Document,
  story: OpStory,
  gap: Gap,
  runProps: InsertedRunProps,
) => {
  const applied = applyDocumentOps(base, [
    {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story, blockId: BLOCK, ...gap },
      text: TYPED,
      runProps,
    },
  ]);
  return applied.isErr() ? applied.error.message : applied.value;
};

const leaf = fc.oneof(
  fc.constantFrom<RunContent>(
    { type: "text", text: "ab" },
    { type: "text", text: "cde" },
  ),
  fc.constantFrom(...RUN_LEAF_SAMPLES),
);

const BOOKMARK_START: ParagraphContent = {
  type: "bookmarkStart",
  id: 7,
  name: "mark",
};
const BOOKMARK_END: ParagraphContent = { type: "bookmarkEnd", id: 7 };

/**
 * Runs of leaves of every kind, some bold, with a bookmark's two ends
 * between them: zero-width leaves inside runs and between them.
 */
const paragraphContent = fc
  .tuple(
    fc.array(
      fc.record({
        content: fc.array(leaf, { minLength: 1, maxLength: 3 }),
        bold: fc.boolean(),
      }),
      { minLength: 1, maxLength: 3 },
    ),
    fc.nat(),
    fc.nat(),
    fc.boolean(),
  )
  .map(([runs, first, second, bookmark]): ParagraphContent[] => {
    const items: ParagraphContent[] = runs.map(({ content, bold }) =>
      bold
        ? { type: "run", formatting: { bold: true }, content }
        : { type: "run", content },
    );
    if (!bookmark) return items;
    const start = first % (items.length + 1);
    const end = start + (second % (items.length + 1 - start));
    return items
      .slice(0, start)
      .concat([BOOKMARK_START], items.slice(start, end), [BOOKMARK_END])
      .concat(items.slice(end));
  })
  // A note holds one automatic reference mark.
  .filter(
    (content) =>
      leafSpans(content).filter(({ node }) => node.type === "noteMarker")
        .length <= 1,
  );

const scenario = paragraphContent.chain((content) => {
  const gaps = gapsOf(content);
  return fc.record({
    content: fc.constant(content),
    gap: fc.constantFrom(...gaps),
    runProps: fc.constantFrom<InsertedRunProps>("inherit", ITALIC),
  });
});

type Scenario =
  typeof scenario extends fc.Arbitrary<infer Value> ? Value : never;

const EXAMPLES: [Scenario][] = [
  // After a run holding only a note's own mark: the text joins that run.
  [
    {
      content: [
        { type: "run", content: [RUN_LEAVES.noteMarker] },
        { type: "run", content: [{ type: "text", text: "ab" }] },
      ],
      gap: { offset: 0, zeroWidthBefore: 1 },
      runProps: "inherit",
    },
  ],
  // Between a run's text and a cached page break it ends with.
  [
    {
      content: [
        {
          type: "run",
          content: [{ type: "text", text: "ab" }, RUN_LEAVES.renderedPageBreak],
        },
        { type: "run", formatting: { bold: true }, content: [RUN_LEAVES.tab] },
      ],
      gap: { offset: 2, zeroWidthBefore: 0 },
      runProps: "inherit",
    },
  ],
  // After a bookmark's end: the text stays outside the bookmark.
  [
    {
      content: [
        BOOKMARK_START,
        { type: "run", content: [{ type: "text", text: "ab" }] },
        BOOKMARK_END,
        { type: "run", formatting: { bold: true }, content: [RUN_LEAVES.tab] },
      ],
      gap: { offset: 2, zeroWidthBefore: 1 },
      runProps: "inherit",
    },
  ],
];

/**
 * Typing at the sampled gap; `misplaced` is the test-only break, which
 * types at the next gap instead.
 */
const lands = (misplaced = false) =>
  fc.property(scenario, ({ content, gap, runProps }) => {
    const { story, document } = placeOf(content);
    const gaps = gapsOf(content);
    const typed = misplaced
      ? (gaps[
          (gaps.findIndex((each) => compareGaps(each, gap) === 0) + 1) %
            gaps.length
        ] ?? gap)
      : gap;
    expect(
      landing(
        document,
        story,
        gap,
        runProps,
        typeAt(document, story, typed, runProps),
      ),
    ).toBeUndefined();
  });

describe("typing at a gap", () => {
  test(
    "an interior insertion changes only the original text leaf",
    () => {
      assertProperty(
        fc.property(
          fc.array(fc.string({ minLength: 2, maxLength: 12 }), {
            minLength: 1,
            maxLength: 5,
          }),
          fc.nat(),
          fc.nat(),
          (texts, leafIndex, interiorOffset) => {
            const index = leafIndex % texts.length;
            const text = texts[index];
            if (text === undefined)
              throw new Error("The selected text leaf exists.");
            const cut = 1 + (interiorOffset % (text.length - 1));
            const offset =
              texts
                .slice(0, index)
                .reduce((total, value) => total + value.length, 0) + cut;
            const content: ParagraphContent[] = [
              {
                type: "run",
                content: texts.map((value) => ({ type: "text", text: value })),
              },
            ];
            const { document, story } = placeOf(content);
            const applied = typeAt(
              document,
              story,
              { offset, zeroWidthBefore: 0 },
              "inherit",
            );
            if (typeof applied === "string") throw new Error(applied);
            expect(paragraphOf(applied.document, story)).toStrictEqual([
              {
                type: "run",
                content: texts.map((value, at) => ({
                  type: "text",
                  text:
                    at === index
                      ? value.slice(0, cut) + TYPED + value.slice(cut)
                      : value,
                })),
              },
            ]);
            const restored = applyDocumentOps(
              applied.document,
              JSON.parse(JSON.stringify(applied.inverse)),
            ).unwrap();
            expect(restored.document).toStrictEqual(document);
          },
        ),
        {
          id: "interior typing preserves original text leaf boundaries",
          numRuns: 200,
        },
      );
    },
    propertyTestTimeout(30_000),
  );

  test(
    "lands exactly at the gap, in the run beside it, whatever leaves are there",
    () => {
      fc.assert(lands(), propertyConfig({ numRuns: 300, examples: EXAMPLES }));
    },
    propertyTestTimeout(60_000),
  );

  test(
    "typing at another gap than the one named is caught",
    () => {
      const result = fc.check(
        lands(true),
        // The break is a counterexample on purpose: report nothing.
        propertyConfig({
          numRuns: 1,
          examples: EXAMPLES,
          reporter: () => undefined,
        }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(30_000),
  );

  test("text after a run holding only a note's own mark joins that run", () => {
    const content: ParagraphContent[] = [
      { type: "run", content: [RUN_LEAVES.noteMarker] },
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ];
    const { story, document } = placeOf(content);
    const applied = typeAt(
      document,
      story,
      { offset: 0, zeroWidthBefore: 1 },
      "inherit",
    );
    if (typeof applied === "string") throw new Error(applied);
    expect(paragraphOf(applied.document, story)).toEqual([
      {
        type: "run",
        content: [RUN_LEAVES.noteMarker, { type: "text", text: TYPED }],
      },
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ]);
    // The oracle catches the text in the run after the mark's.
    expect(
      landing(document, story, { offset: 0, zeroWidthBefore: 1 }, "inherit", {
        document: placeOf([
          { type: "run", content: [RUN_LEAVES.noteMarker] },
          { type: "run", content: [{ type: "text", text: `${TYPED}ab` }] },
        ]).document,
        inverse: [],
      }),
    ).toEqual({ host: "with the leaf after", wanted: "with the leaf before" });
  });
});
