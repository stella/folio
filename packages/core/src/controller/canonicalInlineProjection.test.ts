import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { OP_STORIES, paragraphLogicalText } from "@stll/docx-core/ops";
import type { Document, Paragraph, ParagraphContent, Run, ComplexField } from "../types/document";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { projectCanonicalInline } from "./canonicalInlineProjection";
import { createCanonicalSession } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(30_000));

test("structured field payloads validate text without refusing authored instruction runs", () => {
  const field = {
    type: "complexField",
    instruction: "REF target",
    fieldType: "REF",
    fieldCode: [{ type: "run", content: [{ type: "instrText", text: "REF target" }] }],
    fieldResult: [run("cached")],
  } satisfies ComplexField;
  const source = {
    type: "paragraph",
    paraId: "12345678",
    content: [run("L"), field, run("R")],
  } satisfies Paragraph;
  expect(createCanonicalSession(documentFor(source)).isOk()).toBe(true);
  for (const text of ["\u0001", "\t", "\n"]) {
    field.fieldResult = [run(text)];
    const result = createCanonicalSession(documentFor(source));
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.message).toContain("valid XML characters");
  }
});

const run = (text: string) => ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;
type RangeStart = Extract<ParagraphContent["type"], "bookmarkStart" | `${string}RangeStart`>;
const PAIRS = {
  bookmarkStart: [
    { type: "bookmarkStart", id: 1, name: "target" },
    { type: "bookmarkEnd", id: 1 },
  ],
  commentRangeStart: [
    { type: "commentRangeStart", id: 1 },
    { type: "commentRangeEnd", id: 1 },
  ],
  moveFromRangeStart: [
    { type: "moveFromRangeStart", id: 1, name: "source", author: "Reviewer" },
    { type: "moveFromRangeEnd", id: 1 },
  ],
  moveToRangeStart: [
    { type: "moveToRangeStart", id: 1, name: "destination", author: "Reviewer" },
    { type: "moveToRangeEnd", id: 1 },
  ],
} as const satisfies Record<RangeStart, readonly [ParagraphContent, ParagraphContent]>;

const documentFor = (paragraph: Paragraph) =>
  ({
    package: {
      document: {
        content: [paragraph],
        comments: paragraph.content.flatMap((item) =>
          item.type === "commentRangeStart"
            ? [
                {
                  id: item.id,
                  author: "Reviewer",
                  content: [
                    {
                      type: "paragraph",
                      paraId: (0x20000000 + item.id).toString(16),
                      content: [run("Comment")],
                    } satisfies Paragraph,
                  ],
                },
              ]
            : [],
        ),
      },
    },
  }) satisfies Document;

test("generated collapsed and paired markers retain exact source gap ordinals", () => {
  const check = (pairs: (typeof PAIRS)[RangeStart][]) => {
    const content: ParagraphContent[] = [run("L")];
    const expected = [{ position: 1, zeroWidthBefore: 0 }];
    let position = 1;
    let ordinal = 0;
    for (const [index, pair] of pairs.entries()) {
      content.push({ ...pair[0], id: index + 1 }, { ...pair[1], id: index + 1 });
      if (pair[0].type === "bookmarkStart") {
        expected.push({ position: ++position, zeroWidthBefore: ++ordinal });
        expected.push({ position: ++position, zeroWidthBefore: ++ordinal });
      } else {
        ordinal += 2;
        expected.push({ position: ++position, zeroWidthBefore: ordinal });
      }
    }
    content.push(run("R"));
    const source = { type: "paragraph", paraId: "12345678", content } satisfies Paragraph;
    const document = documentFor(source);
    const native = toProseDoc(document).child(0);
    const mapped = projectCanonicalInline(source, native).unwrap();
    expect(mapped.text).toBe(paragraphLogicalText(source));
    expect(mapped.text).toBe("LR");
    expect(mapped.boundaries).toEqual([
      [{ position: 0, zeroWidthBefore: 0 }],
      expected,
      [{ position: position + 1, zeroWidthBefore: 0 }],
    ]);
    const projection = createCanonicalSession(document).unwrap().projection;
    for (const gap of expected) {
      const address = projection.addressAt(1 + gap.position).unwrap();
      expect(address.offset).toBe(1);
      expect(address.zeroWidthBefore).toBe(gap.zeroWidthBefore);
      expect(projection.positionAt(address).unwrap()).toBe(1 + gap.position);
    }
    for (let missing = 0; missing <= ordinal; missing++) {
      if (expected.some((gap) => gap.zeroWidthBefore === missing)) continue;
      const refused = projection.positionAt({
        story: OP_STORIES.MAIN,
        blockId: source.paraId,
        offset: 1,
        zeroWidthBefore: missing,
      });
      expect(refused.isErr()).toBe(true);
      if (refused.isErr()) expect(refused.error.reason).toBe("refused");
    }
  };
  for (const pair of Object.values(PAIRS)) check([pair]);
  check(Object.values(PAIRS));
  assertProperty(
    fc.property(
      fc.array(fc.constantFrom(...Object.values(PAIRS)), { minLength: 1, maxLength: 6 }),
      check,
    ),
    { seed: -2053065844, numRuns: 40 },
  );
});

test("generated nonempty bookmark and move ranges preserve native gap geometry", () => {
  const pairs = Object.values(PAIRS).filter((pair) => pair[0].type !== "commentRangeStart");
  assertProperty(
    fc.property(fc.constantFrom(...pairs), fc.constantFrom("a", "契", "e\u0301"), (pair, text) => {
      const source = {
        type: "paragraph",
        paraId: "12345678",
        content: [run("L"), pair[0], run(text), pair[1], run("R")],
      } satisfies Paragraph;
      const document = documentFor(source);
      const native = toProseDoc(document).child(0);
      const mapped = projectCanonicalInline(source, native).unwrap();
      expect(mapped.text).toBe(paragraphLogicalText(source));
      expect(native.content.size).toBe(mapped.text.length + 2);
      expect(mapped.boundaries.at(1)).toEqual([
        { position: 1, zeroWidthBefore: 0 },
        { position: 2, zeroWidthBefore: 1 },
      ]);
      expect(mapped.boundaries.at(1 + text.length)).toEqual([
        { position: 2 + text.length, zeroWidthBefore: 0 },
        { position: 3 + text.length, zeroWidthBefore: 1 },
      ]);
      const projection = createCanonicalSession(document).unwrap().projection;
      for (const [offset, gaps] of mapped.boundaries.entries()) {
        for (const gap of gaps) {
          const address = projection.addressAt(1 + gap.position).unwrap();
          expect(address.offset).toBe(offset);
          expect(address.zeroWidthBefore ?? 0).toBe(gap.zeroWidthBefore);
          expect(projection.positionAt(address).unwrap()).toBe(1 + gap.position);
        }
      }
    }),
    { seed: 197, numRuns: 30 },
  );
});

test("source markers with no native boundary receive typed activation refusals", () => {
  const contents = [
    [run("L"), PAIRS.commentRangeStart[0], run("x"), PAIRS.commentRangeStart[1], run("R")],
    [run("L"), PAIRS.bookmarkStart[0], run("R")],
  ] satisfies ParagraphContent[][];
  for (const content of contents) {
    const source = { type: "paragraph", paraId: "12345678", content } satisfies Paragraph;
    const result = createCanonicalSession(documentFor(source));
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.reason).toBe("refused");
  }
});
