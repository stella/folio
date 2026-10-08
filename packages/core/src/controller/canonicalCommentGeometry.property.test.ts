import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { inlineLeafSpans, paragraphLogicalText } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, ParagraphContent, RunContent } from "../types/document";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import { EditorState } from "prosemirror-state";
import { documentShape, shapeArrayBuffer } from "../__tests__/documentShapes";
import { parseShapeDocument, placeSelection } from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { prepareCanonicalCommands } from "./canonicalStructure";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { projectCanonicalInline } from "./canonicalInlineProjection";

setDefaultTimeout(propertyTestTimeout(30_000));

const UNIT_COUNT = 6;
const interval = fc
  .tuple(fc.integer({ min: 0, max: UNIT_COUNT }), fc.integer({ min: 0, max: UNIT_COUNT }))
  .map(([left, right]) => ({ from: Math.min(left, right), to: Math.max(left, right) }));
const UNIT_KINDS = ["text", "tab", "image"] as const;
const REVISIONS = ["live", "insertion", "deletion"] as const;
const unit = fc.record({
  kind: fc.constantFrom(...UNIT_KINDS),
  revision: fc.constantFrom(...REVISIONS),
});
type Interval = { from: number; to: number };
type Unit = { kind: (typeof UNIT_KINDS)[number]; revision: (typeof REVISIONS)[number] };

const sourceUnit = ({ kind, revision }: Unit, index: number): ParagraphContent => {
  let content: RunContent;
  switch (kind) {
    case "text":
      content = { type: "text", text: String.fromCharCode(65 + index) };
      break;
    case "tab":
      content = { type: "tab" };
      break;
    case "image":
      content = {
        type: "drawing",
        image: { type: "image", size: { width: 914400, height: 914400 }, wrap: { type: "inline" } },
      };
      break;
    default:
      return panic(String(kind satisfies never));
  }
  const run = { type: "run", content: [content] } satisfies ParagraphContent;
  return revision === "live"
    ? run
    : {
        type: revision,
        info: { id: 100 + index, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
        content: [run],
      };
};

const documentFor = (source: Paragraph): Document => ({
  package: {
    document: {
      content: [source],
      comments: inlineLeafSpans(source.content).flatMap(({ node }) =>
        node.type === "commentRangeStart"
          ? [
              {
                id: node.id,
                author: "Reviewer",
                content: [
                  {
                    type: "paragraph",
                    paraId: (0x20000000 + node.id).toString(16),
                    content: [{ type: "run", content: [{ type: "text", text: "Comment" }] }],
                  },
                ],
              },
            ]
          : [],
      ),
    },
  },
});

const checkGeometry = (ranges: readonly Interval[], units: readonly Unit[]) => {
  const content: ParagraphContent[] = [];
  const expected: { position: number; zeroWidthBefore: number }[][] = [];
  let position = 0;
  for (let offset = 0; offset <= UNIT_COUNT; offset += 1) {
    let ordinal = 0;
    for (const [index, range] of ranges.entries()) {
      if (range.from < range.to && range.to === offset) {
        content.push({ type: "commentRangeEnd", id: index + 1 });
        ordinal += 1;
      }
    }
    for (const [index, range] of ranges.entries()) {
      if (range.from < range.to && range.from === offset) {
        content.push({ type: "commentRangeStart", id: index + 1 });
        ordinal += 1;
      }
    }
    const gaps = [{ position, zeroWidthBefore: ordinal }];
    for (const [index, range] of ranges.entries()) {
      if (range.from === offset && range.to === offset) {
        content.push(
          { type: "commentRangeStart", id: index + 1 },
          { type: "commentRangeEnd", id: index + 1 },
        );
        ordinal += 2;
        position += 1; // One explicit rangeAnchor carries each empty pair.
        gaps.push({ position, zeroWidthBefore: ordinal });
      }
    }
    expected.push(gaps);
    const item = units.at(offset);
    if (item) {
      content.push(sourceUnit(item, offset));
      position += 1; // Text code units, tabs and inline images each occupy one PM unit.
    }
  }
  const source = { type: "paragraph", paraId: "12345678", content } satisfies Paragraph;
  const document = documentFor(source);
  const unchanged = structuredClone(document);
  const paragraph = toProseDoc(document).child(0);
  // Every visible carrier keeps all enclosing ranges, including leaf atoms whose
  // own child-content mark policy is empty.
  paragraph.forEach((node, nativePosition) => {
    if (node.type.name === "rangeAnchor") return;
    const offset = expected.findIndex((gaps) => gaps.at(-1)?.position === nativePosition);
    expect(offset).toBeGreaterThanOrEqual(0);
    expect(
      node.marks
        .filter((mark) => mark.type.name === "comment")
        .map((mark) => mark.attrs.commentId)
        .toSorted((left, right) => left - right),
    ).toEqual(
      ranges.flatMap((range, index) =>
        range.from <= offset && offset < range.to ? [index + 1] : [],
      ),
    );
  });
  const mapped = projectCanonicalInline({
    source,
    paragraph,
    pairedBookmarkIds: new Set(),
  }).unwrap();
  expect(mapped.text).toBe(paragraphLogicalText(source));
  expect(mapped.boundaries).toEqual(expected);
  // Bind the oracle's ordinals to the operation owner's source leaves.
  const spans = inlineLeafSpans(content);
  for (const [offset, gaps] of expected.entries()) {
    const sourceOrdinal = Math.max(
      0,
      ...spans
        .filter(({ after }) => after.offset === offset)
        .map(({ after }) => after.zeroWidthBefore),
    );
    expect(gaps.at(-1)?.zeroWidthBefore).toBe(sourceOrdinal);
  }
  expect(createCanonicalSession(document).isOk()).toBe(true);
  expect(document).toEqual(unchanged);
};

test("comment mark transitions preserve source geometry across ranges, atoms and revisions", () => {
  const units: Unit[] = [
    { kind: "text", revision: "live" },
    { kind: "image", revision: "insertion" },
    { kind: "tab", revision: "deletion" },
    { kind: "text", revision: "insertion" },
    { kind: "image", revision: "deletion" },
    { kind: "text", revision: "live" },
  ];
  for (const ranges of [
    [
      { from: 0, to: 6 },
      { from: 1, to: 5 },
    ], // Nested.
    [
      { from: 0, to: 4 },
      { from: 2, to: 6 },
    ], // Overlapping.
    [
      { from: 0, to: 3 },
      { from: 3, to: 6 },
    ], // Adjacent.
    [
      { from: 0, to: 6 },
      { from: 2, to: 2 },
      { from: 2, to: 2 },
    ], // Empty inside a range.
  ])
    checkGeometry(ranges, units);
  assertProperty(
    fc.property(
      fc.array(interval, { minLength: 1, maxLength: 5 }),
      fc.array(unit, { minLength: UNIT_COUNT, maxLength: UNIT_COUNT }),
      checkGeometry,
    ),
    { numRuns: 40 },
  );
});

test("missing or unattributed comment mark transitions refuse without changing source", () => {
  const source = {
    type: "paragraph",
    paraId: "12345678",
    content: [
      { type: "commentRangeStart", id: 1 },
      { type: "run", content: [{ type: "text", text: "x" }] },
      { type: "commentRangeEnd", id: 1 },
    ],
  } satisfies Paragraph;
  const unchanged = structuredClone(source);
  const native = toProseDoc(documentFor(source)).child(0);
  for (const marks of [[], [schema.mark("comment", { commentId: 2 })]]) {
    const paragraph = native.type.create(native.attrs, [schema.text("x", marks)]);
    expect(
      projectCanonicalInline({ source, paragraph, pairedBookmarkIds: new Set() }).isErr(),
    ).toBe(true);
    expect(source).toEqual(unchanged);
  }
});

test("comment ranges retain their context across paragraph boundaries", () => {
  const document = {
    package: {
      document: {
        comments: [
          {
            id: 1,
            author: "Reviewer",
            content: [
              {
                type: "paragraph",
                paraId: "20000001",
                content: [{ type: "run", content: [{ type: "text", text: "Comment" }] }],
              },
            ],
          },
        ],
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            content: [
              { type: "commentRangeStart", id: 1 },
              { type: "run", content: [{ type: "text", text: "A" }] },
            ],
          },
          {
            type: "paragraph",
            paraId: "12345679",
            content: [
              { type: "run", content: [{ type: "text", text: "B" }] },
              { type: "commentRangeEnd", id: 1 },
              { type: "commentReference", id: 1 },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  expect(createCanonicalSession(document).isOk()).toBe(true);
});

const commentGeometry = (document: Document) =>
  document.package.document.content.flatMap((paragraph) =>
    paragraph.type !== "paragraph"
      ? []
      : inlineLeafSpans(paragraph.content).flatMap(({ node, before, after }) =>
          node.type === "commentRangeStart" ||
          node.type === "commentRangeEnd" ||
          node.type === "commentReference"
            ? [{ paragraph: paragraph.paraId, type: node.type, id: node.id, before, after }]
            : [],
        ),
  );

test("parsed comments fixture executes canonical formatting and saves its exact comment geometry", async () => {
  const source = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("comments")));
  for (const mode of [{ type: "editing" }, { type: "suggesting", author: "Reviewer" }] as const) {
    const session = createCanonicalSession(source).unwrap();
    session.setMode(mode);
    const state = placeSelection(
      EditorState.create({ doc: session.projection.doc }),
      documentShape("comments").focus,
      "word",
    );
    expect(state).not.toBeNull();
    if (!state) throw new TypeError("Comments fixture lost its word selection");
    const command = singletonManager.requireCommand("toggleBold")();
    const intents = getCanonicalCommandIntents(command, state);
    if (!intents) throw new TypeError("Bold command lost its canonical descriptor");
    const baseline = commentGeometry(session.document);
    const commit = prepareCanonicalCommands(session, state, intents).unwrap();
    expect(publishCanonicalProjection({ session, state, commit }).isOk()).toBe(true);
    expect(commentGeometry(session.document)).toEqual(baseline);
    const reopened = await parseShapeDocument(new Uint8Array(await createDocx(session.document)));
    expect(commentGeometry(reopened)).toEqual(baseline);
    expect(createCanonicalSession(reopened).isOk()).toBe(true);
  }
});
