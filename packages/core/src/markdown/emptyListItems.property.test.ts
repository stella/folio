import { expect, test } from "bun:test";
import fc from "fast-check";

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
