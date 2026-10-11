import { expect, test } from "bun:test";
import fc from "fast-check";
import { lexMarkdown } from "../../../docx-core/src/markdown/lexer";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { BlockContent, ListRendering, Paragraph } from "../types/document";
import { fromMarkdown } from "./fromMarkdown";
import { toMarkdown } from "./index";

type EmptyListCase = {
  level: number;
  parentMarkers: boolean[];
  isBullet: boolean;
  text: string;
  hasSibling: boolean;
};

const listRendering = (level: number, isBullet: boolean): ListRendering => ({
  marker: isBullet ? "-" : "1.",
  level,
  numId: level + 1,
  isBullet,
  ...(isBullet ? {} : { numFmt: "decimal" }),
});

const paragraph = (text: string, list: ListRendering): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
  listRendering: list,
});

const paragraphText = ({ content }: Paragraph): string =>
  content
    .flatMap((item) =>
      item.type === "run"
        ? item.content.flatMap((part) => (part.type === "text" ? [part.text] : []))
        : [],
    )
    .join("");

const emptyListCaseArbitrary = fc.integer({ min: 1, max: 8 }).chain((level) =>
  fc.record({
    level: fc.constant(level),
    parentMarkers: fc.array(fc.boolean(), { minLength: level, maxLength: level }),
    isBullet: fc.boolean(),
    text: fc.constantFrom("", " ", "\t"),
    hasSibling: fc.boolean(),
  }),
);

const verifyRoundTrip = ({
  level,
  parentMarkers,
  isBullet,
  text,
  hasSibling,
}: EmptyListCase): void => {
  const source: BlockContent[] = parentMarkers.map((isParentBullet, parentLevel) =>
    paragraph(`Parent ${parentLevel}`, listRendering(parentLevel, isParentBullet)),
  );
  source.push(paragraph(text, listRendering(level, isBullet)));
  if (hasSibling) source.push(paragraph("Sibling", listRendering(level, isBullet)));

  const markdown = toMarkdown({ package: { document: { content: source } } });
  const imported = fromMarkdown(markdown);
  const actual = imported.package.document.content.map((block) => {
    expect(block.type).toBe("paragraph");
    if (block.type !== "paragraph") throw new Error(`Unexpected block ${block.type}`);
    return {
      text: paragraphText(block),
      level: block.listRendering?.level,
      isBullet: block.listRendering?.isBullet,
    };
  });
  const expected = parentMarkers.map((isParentBullet, parentLevel) => ({
    text: `Parent ${parentLevel}`,
    level: parentLevel,
    isBullet: isParentBullet,
  }));
  expected.push({ text: text.trim(), level, isBullet });
  if (hasSibling) expected.push({ text: "Sibling", level, isBullet });

  expect(actual).toEqual(expected);
};

// Earlier list-context fixtures only used nonempty children. Exercise every
// native marker and supported depth with empty and whitespace-only children.
const REGRESSION_EXAMPLES: EmptyListCase[] = Array.from({ length: 8 }, (_, index) =>
  [true, false].flatMap((isBullet) =>
    ["", " ", "\t"].flatMap((text) =>
      [true, false].map((hasSibling) => ({
        level: index + 1,
        parentMarkers: Array.from(
          { length: index + 1 },
          (_parent, parentLevel) => parentLevel % 2 === 0,
        ),
        isBullet,
        text,
        hasSibling,
      })),
    ),
  ),
).flat();

test(
  "empty nested list paragraphs survive Markdown export and import",
  async () => {
    await assertProperty(fc.property(emptyListCaseArbitrary, verifyRoundTrip), {
      id: "empty nested list paragraphs survive Markdown export and import",
      numRuns: 30,
      examples: REGRESSION_EXAMPLES.map((example) => [example]),
    });
  },
  propertyTestTimeout(30_000),
);

const markerArbitrary = fc.constantFrom("-", "1.", "a.", "i)", "1.2.", "=", "---");
const mixedItemArbitrary = fc.record({
  level: fc.integer({ min: 0, max: 8 }),
  marker: markerArbitrary,
});
type MixedItem = fc.ArbitraryValue<typeof mixedItemArbitrary>;

const verifyMixedKinds = (items: MixedItem[]): void => {
  const source = items.map(({ level, marker }, index) =>
    paragraph(`Item ${index}`, {
      marker,
      level,
      numId: index + 1,
      isBullet: marker === "-",
    }),
  );
  const markdown = toMarkdown({ package: { document: { content: source } } });
  // Inspect the same grammar used by the importer before code blocks become
  // ordinary model paragraphs with monospace runs.
  const tokens = lexMarkdown(markdown);
  const pending = [...tokens];
  for (let token = pending.pop(); token !== undefined; token = pending.pop()) {
    expect(
      token.type === "list" ||
        token.type === "text" ||
        token.type === "paragraph" ||
        token.type === "space",
    ).toBe(true);
    if (token.type === "list" && "items" in token) {
      for (const item of token.items) pending.push(...item.tokens);
    }
  }
  const imported = fromMarkdown(markdown).package.document.content;
  expect(
    imported.map((block) => {
      expect(block.type).toBe("paragraph");
      if (block.type !== "paragraph") throw new Error(`Unexpected block ${block.type}`);
      return {
        text: paragraphText(block),
        kind: block.listRendering === undefined ? "paragraph" : "listItem",
      };
    }),
  ).toEqual(
    items.map(({ marker }, index) => {
      const native = marker === "-" || marker === "1.";
      return {
        text: native ? `Item ${index}` : `${marker} Item ${index}`,
        kind: native ? "listItem" : "paragraph",
      };
    }),
  );
};

const mixedExamples: MixedItem[][] = Array.from({ length: 9 }, (_, level) =>
  ["a.", "i)", "1.2.", "=", "---"].flatMap((marker) => {
    const ancestors = Array.from({ length: level + 1 }, (_ancestor, depth) => ({
      level: depth,
      marker: depth % 2 === 0 ? "-" : "1.",
    }));
    const custom = { level, marker };
    return [
      [...ancestors, custom, ...ancestors],
      [custom, ...ancestors, custom],
    ];
  }),
).flat();

test(
  "mixed native and custom list labels preserve Markdown block kinds",
  async () => {
    await assertProperty(
      fc.property(fc.array(mixedItemArbitrary, { minLength: 1, maxLength: 24 }), verifyMixedKinds),
      {
        id: "mixed native and custom list labels preserve Markdown block kinds",
        numRuns: 100,
        examples: mixedExamples.map((example) => [example]),
      },
    );
  },
  propertyTestTimeout(30_000),
);
