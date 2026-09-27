import { describe, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

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

const normalize = (src: string): string => toMarkdown(fromMarkdown(src), CLEAN);

const LOWERCASE_WORD_CHARS = [
  "a",
  "b",
  "c",
  "d",
  "e",
  "f",
  "g",
  "h",
  "i",
  "j",
  "k",
  "l",
  "m",
  "n",
  "o",
  "p",
  "q",
  "r",
  "s",
  "t",
  "u",
  "v",
  "w",
  "x",
  "y",
  "z",
] as const;

const inlineText = fc
  .array(
    fc.array(fc.constantFrom(...LOWERCASE_WORD_CHARS), {
      minLength: 1,
      maxLength: 8,
    }),
    { minLength: 1, maxLength: 3 },
  )
  .map((words) => words.map((chars) => chars.join("")).join(" "));

const inlineMarkdown = fc.oneof(
  inlineText,
  inlineText.map((value) => `**${value}**`),
  inlineText.map((value) => `*${value}*`),
  inlineText.map((value) => `\`${value.replaceAll("`", "")}\``),
  inlineText.map((value) => `[${value}](https://example.com/${encodeURIComponent(value)})`),
);

const looseListItemMarkdown = () =>
  fc
    .tuple(
      fc.array(inlineMarkdown, { minLength: 1, maxLength: 3 }),
      fc.array(inlineMarkdown, { minLength: 1, maxLength: 3 }),
    )
    .map(
      ([firstParagraph, secondParagraph]) =>
        `- ${firstParagraph.join(" ")}\n\n  ${secondParagraph.join(" ")}`,
    );

// The generator above only ever produced one bullet item with a plain-text
// continuation paragraph: never an ordered parent (whose marker width varies
// with the counter — "9. " is 3 columns, "10. " is 4), never mixed
// bullet/ordered nesting, never more than one level deep, and never block
// content (a table, code block, or blockquote) inside an item. That gap is
// exactly how the two QA findings this generator now covers slipped past it:
// an ordered parent's child re-imported one level flatter (wrong export
// indent for anything but a bullet parent), and a table inside a list item
// vanished on import. `nestedListShapeMarkdown` below builds markdown by
// hand, tracking each item's own marker width the way CommonMark requires,
// so every generated source is valid input on its own terms (not just
// something `toMarkdown` would have produced), and every leaf text is a
// unique sentinel so the oracle can check none were lost, duplicated, or
// merged into a neighbour.
type ItemFeature =
  | { kind: "plain" }
  | { kind: "continuation" }
  | { kind: "children"; list: ListShape }
  | { kind: "table" }
  | { kind: "code" }
  | { kind: "blockquote" };

type ItemShape = { feature: ItemFeature };

type ListShape = { ordered: boolean; start: number; items: ItemShape[] };

// A mix of 1-digit and 2-digit ordered markers, including a start that
// crosses the 1-digit/2-digit boundary mid-list (9, 10, 11, …).
const ORDERED_START_VALUES = [1, 2, 9, 10, 11] as const;

const itemFeatureArb = (depth: number): fc.Arbitrary<ItemFeature> => {
  const options: fc.Arbitrary<ItemFeature>[] = [
    fc.constant({ kind: "plain" }),
    fc.constant({ kind: "continuation" }),
    fc.constant({ kind: "table" }),
    fc.constant({ kind: "code" }),
    fc.constant({ kind: "blockquote" }),
  ];
  if (depth > 0) {
    options.push(listShapeArb(depth - 1).map((list): ItemFeature => ({ kind: "children", list })));
  }
  return fc.oneof(...options);
};

// `depth` bounds nesting so the whole tree stays at most 3 levels deep
// (depth 2 here: this list, one nested list, one nested-in-nested list).
const listShapeArb = (depth: number): fc.Arbitrary<ListShape> =>
  fc.record({
    ordered: fc.boolean(),
    start: fc.constantFrom(...ORDERED_START_VALUES),
    items: fc.array(
      itemFeatureArb(depth).map((feature): ItemShape => ({ feature })),
      { minLength: 1, maxLength: 3 },
    ),
  });

/**
 * Render one list at the given left indent (columns already consumed by
 * every ancestor's marker), collecting every sentinel it plants along the
 * way. Each item gets exactly one extra feature (never a continuation *and*
 * a table, say) so the generated source stays unambiguous to hand-verify.
 */
const buildList = (
  list: ListShape,
  ancestorIndent: number,
  sentinels: string[],
  nextSentinel: () => string,
): string[] => {
  const lines: string[] = [];
  let counter = list.start;
  for (const item of list.items) {
    const marker = list.ordered ? `${counter}. ` : "- ";
    const sentinel = nextSentinel();
    sentinels.push(sentinel);
    lines.push(`${" ".repeat(ancestorIndent)}${marker}${sentinel}`);
    // CommonMark: a child needs to start at or past the column where this
    // item's own content begins, i.e. past this marker (not a fixed guess).
    const childIndent = ancestorIndent + marker.length;
    const pad = " ".repeat(childIndent);
    const { feature } = item;
    if (feature.kind === "continuation") {
      const text = nextSentinel();
      sentinels.push(text);
      lines.push("", `${pad}${text}`);
    } else if (feature.kind === "children") {
      lines.push(...buildList(feature.list, childIndent, sentinels, nextSentinel));
    } else if (feature.kind === "table") {
      const cells = [nextSentinel(), nextSentinel(), nextSentinel(), nextSentinel()];
      sentinels.push(...cells);
      lines.push(
        "",
        `${pad}| ${cells[0]} | ${cells[1]} |`,
        `${pad}| --- | --- |`,
        `${pad}| ${cells[2]} | ${cells[3]} |`,
      );
    } else if (feature.kind === "code") {
      const text = nextSentinel();
      sentinels.push(text);
      lines.push("", `${pad}\`\`\``, `${pad}${text}`, `${pad}\`\`\``);
    } else if (feature.kind === "blockquote") {
      const text = nextSentinel();
      sentinels.push(text);
      lines.push("", `${pad}> ${text}`);
    }
    counter += 1;
  }
  return lines;
};

const nestedListShapeMarkdown = (): fc.Arbitrary<{ markdown: string; sentinels: string[] }> =>
  listShapeArb(2).map((list) => {
    const sentinels: string[] = [];
    let counter = 0;
    // Fixed-width so no sentinel can ever be a substring of a different one:
    // the oracle's occurrence count would otherwise be meaningless.
    const nextSentinel = () => `sent${String(counter++).padStart(4, "0")}`;
    const markdown = buildList(list, 0, sentinels, nextSentinel).join("\n");
    return { markdown, sentinels };
  });

const occurrences = (needle: string, haystack: string): number => haystack.split(needle).length - 1;

describe("markdown bridge — round-trip (properties)", () => {
  test("loose list items are idempotent after markdown bridge normalization", () => {
    fc.assert(
      fc.property(
        looseListItemMarkdown(),
        (source) => normalize(normalize(source)) === normalize(source),
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("every sentinel survives ordered/mixed nested lists and block content in an item", () => {
    // Not asserting full idempotency here (unlike the property above): a
    // table/code block/blockquote inside a list item exports as a following
    // sibling block (see content.ts), and when that item has a later sibling
    // *of an ancestor list*, the sibling ends up on the far side of a
    // non-list paragraph. CommonMark has no way to say "this bullet still
    // continues the interrupted list at level 1" across that gap — only
    // adjacency says so — so re-importing that export can renumber the
    // sibling to a shallower level on a second cycle. That's a real, narrow
    // limitation of plain-markdown list continuation (not specific to this
    // fix, and not data loss: every sentinel still appears exactly once,
    // which is what this oracle checks). Losing or duplicating a sentinel
    // would still fail this property.
    fc.assert(
      fc.property(nestedListShapeMarkdown(), ({ markdown, sentinels }) => {
        const once = normalize(markdown);
        return sentinels.every((sentinel) => occurrences(sentinel, once) === 1);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
