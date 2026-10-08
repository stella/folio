import { expect, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import type { Document, Paragraph, ParagraphContent, Run } from "../../model/document";
import { assertProperty } from "../../../../../test/property-testing";
import { RANGE_ANCHOR_FIXTURE_FACTORIES } from "../../../typecheck/range-anchor-fixtures.typecheck";
import { applyDocumentOps } from "../apply";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  paragraphVisibleText,
  type EditorIntent,
} from "../editorIntent";
import { leafSpans, isParagraphContent } from "../leaves";
import { isRangeBoundary, paragraphLogicalText } from "../offsets";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const run = (text: string) => ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;

test("moving a selected bookmark transfers its pair once and restores exactly on undo", () => {
  const selected = [
    { type: "bookmarkStart", id: 11, name: "target", displacedByCustomXml: "next" },
    run("x"),
    { type: "bookmarkEnd", id: 11, displacedByCustomXml: "prev" },
  ] as const;
  const paragraph = {
    type: "paragraph",
    paraId: "12345678",
    content: [run("L"), ...selected, run("R")],
  } satisfies Paragraph;
  const document = { package: { document: { content: [paragraph] } } } satisfies Document;
  const at = { story: OP_STORIES.MAIN, blockId: paragraph.paraId } as const;
  const intent = {
    type: "moveFragment",
    from: { ...at, offset: 1, zeroWidthBefore: 0 },
    to: { ...at, offset: 2, zeroWidthBefore: 1 },
    target: { ...at, offset: 3, zeroWidthBefore: 0 },
    paragraphs: [{ type: "paragraph", content: [...selected] }],
    openStart: 1,
    openEnd: 1,
  } as const satisfies EditorIntent;
  const allocation = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: allocation.newIds },
  }).unwrap();
  const applied = applyDocumentOps(document, compiled.ops).unwrap();
  const moved = applied.document.package.document.content.at(0);
  if (moved?.type !== "paragraph") panic("Move lost its paragraph.");
  expect(paragraphVisibleText(moved)).toBe("LRx");
  const markers = leafSpans(moved.content).filter(
    ({ node }) =>
      isParagraphContent(node) && (node.type === "bookmarkStart" || node.type === "bookmarkEnd"),
  );
  expect(markers).toHaveLength(2);
  expect(markers.map(({ before }) => before.offset)).toEqual([2, 3]);
  const start = markers.at(0)?.node;
  if (start?.type !== "bookmarkStart") panic("Move lost its opening bookmark.");
  expect(start.name).toBe("target");
  expect(start.displacedByCustomXml).toBe("next");
  const end = markers.at(1)?.node;
  if (end?.type !== "bookmarkEnd") panic("Move lost its closing bookmark.");
  expect(end.displacedByCustomXml).toBe("prev");
  expect(end.id).toBe(start.id);
  expect(applyDocumentOps(applied.document, applied.inverse).unwrap().document).toStrictEqual(
    document,
  );
});

type RangeFactory =
  (typeof RANGE_ANCHOR_FIXTURE_FACTORIES)[keyof typeof RANGE_ANCHOR_FIXTURE_FACTORIES];
type TransferArgs = {
  factories: readonly RangeFactory[];
  text: string;
  target: "before" | "after" | "inside";
};

const independentVisibleText = (items: readonly ParagraphContent[]): string =>
  items
    .map((item) => {
      switch (item.type) {
        case "run":
          return item.content.map((child) => (child.type === "text" ? child.text : "")).join("");
        case "insertion":
        case "moveTo":
        case "inlineWrapper":
          return independentVisibleText(item.content);
        case "hyperlink":
          return independentVisibleText(item.children);
        default:
          return "";
      }
    })
    .join("");

const paragraphOf = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Range transfer lost its paragraph.");
  return paragraph;
};
const markersOf = (document: Document) =>
  leafSpans(paragraphOf(document).content).filter(
    ({ node }) => isParagraphContent(node) && isRangeBoundary(node),
  );

test("generated mixed range transfers retain one original pair through suggested review and history", () => {
  const check = ({ factories, text, target }: TransferArgs) => {
    const pairs = factories.map((factory) => factory());
    const starts = pairs.map((pair) => pair[0]);
    const ends = pairs.toReversed().map((pair) => pair[1]);
    const sourceMarkers = [...starts, ...ends];
    const selected = [...starts, ...(text === "" ? [] : [run(text)]), ...ends];
    const paragraph = {
      type: "paragraph",
      paraId: "12345678",
      content: [run("L"), ...selected, run("R")],
    } satisfies Paragraph;
    const comments = starts.flatMap((start) =>
      start.type === "commentRangeStart"
        ? [
            {
              id: start.id,
              author: "Reviewer",
              content: [
                {
                  type: "paragraph",
                  paraId: "22345678",
                  content: [run("Comment")],
                } satisfies Paragraph,
              ],
            },
          ]
        : [],
    );
    const document = {
      package: { document: { content: [paragraph], comments } },
    } satisfies Document;
    const original = structuredClone(document);
    const at = { story: OP_STORIES.MAIN, blockId: paragraph.paraId } as const;
    const insideOffset = 1 + Math.floor(text.length / 2);
    const policies = {
      before: { offset: 0, text: `${text}LR`, movedOffset: 0, sourceOffset: 1 + text.length },
      after: { offset: 2 + text.length, text: `LR${text}`, movedOffset: 2, sourceOffset: 1 },
      inside: { offset: insideOffset, text: `L${text}R`, movedOffset: 1, sourceOffset: 1 },
    } satisfies Record<
      TransferArgs["target"],
      { offset: number; text: string; movedOffset: number; sourceOffset: number }
    >;
    const destination = policies[target];
    const intent = {
      type: "moveFragment",
      from: { ...at, offset: 1, zeroWidthBefore: 0 },
      to: {
        ...at,
        offset: 1 + text.length,
        zeroWidthBefore: text === "" ? 2 * pairs.length : pairs.length,
      },
      target: {
        ...at,
        offset: destination.offset,
        zeroWidthBefore: target === "inside" && insideOffset === 1 ? pairs.length : 0,
      },
      paragraphs: [{ type: "paragraph", content: selected }],
      openStart: 1,
      openEnd: 1,
    } as const satisfies EditorIntent;
    const allocation = allocateEditorIntentIds(document, intent);
    for (const mode of [
      { type: "editing", newIds: allocation.newIds },
      {
        type: "suggesting",
        revision: { id: allocation.revisionId, author: "Reviewer" },
        newIds: allocation.newIds,
      },
    ] as const) {
      const compiled = compileEditorIntent(document, { intent, mode }).unwrap();
      expect(document).toStrictEqual(original);
      if (target === "inside") expect(compiled.ops).toEqual([]);
      const applied = applyDocumentOps(document, compiled.ops).unwrap();
      const moved = paragraphOf(applied.document);
      expect(independentVisibleText(moved.content)).toBe(destination.text);
      expect(paragraphVisibleText(moved)).toBe(destination.text);
      if (mode.type === "editing") expect(paragraphLogicalText(moved)).toBe(destination.text);
      const assertMarkers = (current: Document, startOffset: number, endOffset: number) => {
        const markers = markersOf(current);
        expect(markers).toHaveLength(sourceMarkers.length);
        for (const [index, marker] of markers.entries()) {
          const expected = sourceMarkers.at(index);
          if (!expected || !("id" in marker.node)) panic("Range transfer lost a paired identity.");
          expect({ ...marker.node, id: expected.id }).toEqual(expected);
          if (
            mode.type === "suggesting" ||
            target === "inside" ||
            expected.type === "commentRangeStart" ||
            expected.type === "commentRangeEnd"
          )
            expect(marker.node.id).toBe(expected.id);
          const partner = markers.at(markers.length - index - 1);
          if (!partner || !("id" in partner.node))
            panic("Range transfer lost its matching endpoint.");
          expect(marker.node.id).toBe(partner.node.id);
          expect(marker.before.offset).toBe(index < pairs.length ? startOffset : endOffset);
          const ordinal = text === "" || startOffset === endOffset ? index : index % pairs.length;
          expect(marker.before.zeroWidthBefore).toBe(ordinal);
          expect(marker.after.zeroWidthBefore).toBe(ordinal + 1);
        }
      };
      const startOffset =
        mode.type === "editing" ? destination.movedOffset : destination.sourceOffset;
      assertMarkers(applied.document, startOffset, startOffset + text.length);
      if (mode.type === "suggesting")
        for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT]) {
          const reviewed =
            applied.revisions.length === 0
              ? applied.document
              : applyDocumentOps(applied.document, [
                  {
                    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                    story: OP_STORIES.MAIN,
                    revisionIds: applied.revisions,
                    decision,
                  },
                ]).unwrap().document;
          const originalText = decision === REVISION_DECISIONS.REJECT || target === "inside";
          const expectedText = originalText ? `L${text}R` : destination.text;
          expect(independentVisibleText(paragraphOf(reviewed).content)).toBe(expectedText);
          expect(paragraphVisibleText(paragraphOf(reviewed))).toBe(expectedText);
          expect(paragraphLogicalText(paragraphOf(reviewed))).toBe(expectedText);
          const sourceOffset = originalText ? 1 : destination.sourceOffset;
          assertMarkers(reviewed, sourceOffset, sourceOffset + (originalText ? text.length : 0));
        }
      if (target === "inside") expect(applied.document).toStrictEqual(original);
      const undo = applyDocumentOps(applied.document, applied.inverse).unwrap();
      expect(undo.document).toStrictEqual(original);
      expect(applyDocumentOps(undo.document, undo.inverse).unwrap().document).toStrictEqual(
        applied.document,
      );
    }
  };
  const factories = Object.values(RANGE_ANCHOR_FIXTURE_FACTORIES);
  for (const selection of [
    [],
    ...factories.map((factory) => [factory]),
    factories,
    factories.toReversed(),
  ])
    for (const text of ["", "x", "e\u0301"])
      for (const target of ["before", "after", "inside"] as const)
        check({ factories: selection, text, target });
  assertProperty(
    fc.property(
      fc.record({
        factories: fc.shuffledSubarray(factories),
        text: fc.constantFrom("", "x", "e\u0301"),
        target: fc.constantFrom("before", "after", "inside"),
      }),
      check,
    ),
    { seed: 197, numRuns: 60 },
  );
});
