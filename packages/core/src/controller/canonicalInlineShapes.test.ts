import { normalizeNoteOccurrenceIds } from "../../../../test/note-occurrence-oracle";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import {
  findStoryBody,
  normalizeForOps,
  OP_STORIES,
  paragraphLogicalText,
  validateOpsDocument,
} from "@stll/docx-core/ops";
import { schema } from "../prosemirror/schema";
import {
  footnoteToProseDoc,
  toProseDoc,
  collectPairedBookmarkIds,
} from "../prosemirror/conversion/toProseDoc";
import type { ParagraphContent, RunContent } from "../types/document";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { splitsSurrogatePair } from "../ai-edits/character-boundaries";
import { CanonicalSessionError, createCanonicalSession } from "./canonicalSession";
import { projectCanonicalInline } from "./canonicalInlineProjection";
import {
  CANONICAL_PARAGRAPH_SHAPE_FACTORIES,
  CANONICAL_RUN_SHAPE_FACTORIES,
  canonicalInlineShapeDocument,
  type CanonicalInlineShapeFixture,
} from "../../typecheck/canonical-inline-shapes.typecheck";

setDefaultTimeout(propertyTestTimeout(60_000));

const fixturesFor = (text: string) => [
  ...Object.values(CANONICAL_PARAGRAPH_SHAPE_FACTORIES)
    .map((factory) => factory(text))
    .flat(),
  ...Object.values(CANONICAL_RUN_SHAPE_FACTORIES)
    .map((factory) => factory(text))
    .flat(),
];
const refusalExpected = ({
  item,
  variant,
}: CanonicalInlineShapeFixture<ParagraphContent | RunContent>) =>
  item.type === "inlineSdt" ||
  item.type === "fieldChar" ||
  item.type === "instrText" ||
  item.type === "renderedPageBreak" ||
  ((item.type === "commentRangeStart" || item.type === "commentRangeEnd") && variant === "filled");

const checkShape = (fixture: CanonicalInlineShapeFixture<ParagraphContent | RunContent>) => {
  const source = normalizeForOps(canonicalInlineShapeDocument(fixture));
  expect(validateOpsDocument(source).isOk()).toBe(true);
  const story =
    fixture.story === "main"
      ? OP_STORIES.MAIN
      : { kind: fixture.story, id: fixture.story === "footnote" ? 1 : 2 };
  const body = findStoryBody(source, story);
  expect(body).toBeDefined();
  if (body === undefined) panic("Fixture story is absent");
  const legacy = fixture.story === "main" ? toProseDoc(source) : footnoteToProseDoc(body.content);
  legacy.check();
  const activated = createCanonicalSession(source);
  if (refusalExpected(fixture)) {
    expect(activated.isErr()).toBe(true);
    if (activated.isErr()) {
      expect(activated.error).toBeInstanceOf(CanonicalSessionError);
      expect(activated.error.reason).toBe("refused");
      expect(activated.error.gap).toBe(CANONICAL_GAP.dispatch);
      expect(activated.error.message).not.toBe("The projection changed paragraph structure.");
      expect(activated.error.message.length).toBeGreaterThan(25);
    }
    return;
  }
  if (activated.isErr())
    panic(
      `Unexpected refusal for ${fixture.item.type}/${fixture.variant}: ${activated.error.message}`,
    );
  const session = activated.value;
  const projection = session.projectStory(story).unwrap();
  expect(normalizeNoteOccurrenceIds(projection.doc.toJSON())).toEqual(
    normalizeNoteOccurrenceIds(legacy.toJSON()),
  );
  expect(schema.nodeFromJSON(projection.doc.toJSON()).eq(projection.doc)).toBe(true);
  let start = 1;
  for (const paragraph of body.content) {
    if (paragraph.type !== "paragraph") panic("Fixture contains a nonparagraph");
    const native = projection.doc.nodeAt(start - 1);
    if (native === null) panic("Fixture native paragraph is absent");
    const mapped = projectCanonicalInline({
      source: paragraph,
      paragraph: native,
      pairedBookmarkIds: collectPairedBookmarkIds(body.content),
    }).unwrap();
    expect(mapped.text).toBe(paragraphLogicalText(paragraph));
    for (let position = 0; position <= native.content.size; position += 1) {
      if (mapped.boundaries.some((gaps) => gaps.some((gap) => gap.position === position))) continue;
      expect(projection.addressAt(start + position).isErr()).toBe(true);
    }
    for (const [offset, gaps] of mapped.boundaries.entries()) {
      expect(gaps.length).toBeGreaterThan(0);
      if (splitsSurrogatePair(mapped.text, offset)) continue;
      for (const gap of gaps) {
        const position = start + gap.position;
        const address = projection.addressAt(position).unwrap();
        expect(address.offset).toBe(offset);
        expect(address.zeroWidthBefore ?? 0).toBe(gap.zeroWidthBefore);
        expect(projection.positionAt(address).unwrap()).toBe(position);
      }
    }
    start += native.nodeSize;
  }
};

test("generated source union shapes activate with exact logical gaps or explain a typed refusal", () => {
  const exercise = (text: string) => {
    // Note references share one model discriminator branch; pin each factory's producer kind.
    for (const [kind, factory] of Object.entries(CANONICAL_RUN_SHAPE_FACTORIES))
      for (const fixture of factory(text)) expect(fixture.item.type).toBe(kind);
    const fixtures = fixturesFor(text);
    expect(new Set(fixtures.map(({ item }) => item.type))).toEqual(
      new Set([
        ...Object.keys(CANONICAL_PARAGRAPH_SHAPE_FACTORIES),
        ...Object.keys(CANONICAL_RUN_SHAPE_FACTORIES),
      ]),
    );
    for (const fixture of fixtures) checkShape(fixture);
  };
  for (const text of ["\u00a0", "e\u0301", "契約", "😀"]) exercise(text);
  assertProperty(
    fc.property(
      fc
        .array(fc.constantFrom("a", "\u00a0", "e\u0301", "契", "約", "😀"), {
          minLength: 1,
          maxLength: 8,
        })
        .map((parts) => parts.join("")),
      exercise,
    ),
    { seed: 197, numRuns: 15 },
  );
});

test("structured field caches keep text safety checks despite atomic projection", () => {
  for (const text of ["bad\u0000", "bad\ud800", "bad\t", "bad\n", "bad\r"]) {
    for (const fixture of [
      ...CANONICAL_PARAGRAPH_SHAPE_FACTORIES.simpleField(text),
      ...CANONICAL_PARAGRAPH_SHAPE_FACTORIES.complexField(text),
    ]) {
      const refused = createCanonicalSession(canonicalInlineShapeDocument(fixture));
      expect(refused.isErr()).toBe(true);
      if (refused.isErr()) expect(refused.error.reason).toBe("refused");
    }
  }
});

test("input inside retained deletion maps through preceding zero-width source markers", () => {
  assertProperty(
    fc.property(fc.integer({ min: 2, max: 8 }), (length) => {
      const fixture = CANONICAL_PARAGRAPH_SHAPE_FACTORIES.bookmarkStart().at(0);
      if (!fixture) panic("Bookmark fixture is absent");
      const [start, end] = fixture.content;
      if (!start || !end) panic("Bookmark fixture lost its paired markers");
      const document = canonicalInlineShapeDocument({
        story: "main",
        content: [
          start,
          {
            type: "deletion",
            info: { id: 42, author: "Other" },
            content: [{ type: "run", content: [{ type: "text", text: "a".repeat(length) }] }],
          },
          end,
        ],
      });
      const projection = createCanonicalSession(document).unwrap().projection;
      for (let interior = 1; interior < length; interior += 1) {
        const address = projection.inputAddressAt(3 + interior).unwrap();
        expect(address.offset).toBe(1 + length);
        expect(address.zeroWidthBefore).toBe(0);
        expect(projection.positionAt(address).unwrap()).toBe(3 + length);
      }
    }),
    { seed: 197, numRuns: 12 },
  );
});
