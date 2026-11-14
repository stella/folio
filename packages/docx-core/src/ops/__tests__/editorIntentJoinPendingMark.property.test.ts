/**
 * A direct paragraph join seen in the editing view.
 *
 * A paragraph whose mark is pending deletion (deleted or moved away) shows
 * as one paragraph with the paragraph after it. Joining two physical
 * paragraphs directly must give the editing view the editor shows for that
 * join, whichever paragraph the join keeps:
 *
 * - when the two are already one paragraph in the editing view, the view is
 *   unchanged, words and paragraph properties;
 * - otherwise the two paragraphs they belong to become one, with the first's
 *   words before the second's, and the first's paragraph properties, or the
 *   second's when the first shows no words;
 * - undoing the join gives the document back exactly.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { assertExactModel } from "../../../../../test/exactModel";
import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphFormatting,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { sameBlockList, storyParagraphs } from "../blocks";
import {
  compileEditorIntent,
  type EditorIntent,
  editorParagraphGroups,
} from "../editorIntent";
import { paragraphPropertiesOf } from "../review";
import { OP_STORIES } from "../types";
import {
  documentArbitrary,
  reviewDocumentArbitrary,
} from "./documentArbitraries";

const SURVIVORS = ["first", "second"] as const;
type Survivor = (typeof SURVIVORS)[number];

const PENDING_MARKS = ["del", "moveFrom"] as const;
type PendingMark = (typeof PENDING_MARKS)[number];

/** Revision ids above every generated one, for the marks a case adds. */
const ADDED_REVISION_IDS = 990_000;

/** Each editing-view paragraph's words and the paragraph properties it shows. */
const editingView = (document: Document) =>
  editorParagraphGroups(document, OP_STORIES.MAIN).map(
    ({ paragraphs, text }) => ({
      text,
      properties: paragraphPropertiesOf(paragraphs.at(-1)?.formatting) ?? {},
    }),
  );

/** The editing view the join of `blockId` and `nextBlockId` must leave. */
const joinedView = (
  document: Document,
  blockId: string,
  nextBlockId: string,
) => {
  const groups = editorParagraphGroups(document, OP_STORIES.MAIN);
  const view = editingView(document);
  const holding = (id: string) =>
    groups.findIndex(({ paragraphs }) =>
      paragraphs.some(({ paraId }) => paraId === id),
    );
  const first = holding(blockId);
  const following = holding(nextBlockId);
  const firstView = view[first];
  const followingView = view[following];
  if (firstView === undefined || followingView === undefined)
    throw new Error("A joined paragraph is in no editing-view paragraph.");
  if (first === following) return view;
  return [
    ...view.slice(0, first),
    {
      text: firstView.text + followingView.text,
      properties:
        firstView.text === "" ? followingView.properties : firstView.properties,
    },
    ...view.slice(following + 1),
  ];
};

type JoinOutcome = "joined" | "refused";

/** Join `blockId` and `nextBlockId` directly, keeping `survivor`, and check the laws. */
const joinLaws = (
  document: Document,
  blockId: string,
  nextBlockId: string,
  survivor: Survivor,
): JoinOutcome => {
  const intent = {
    type: "joinParagraphs",
    story: OP_STORIES.MAIN,
    blockId,
    nextBlockId,
    survivor,
  } as const satisfies EditorIntent;
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing" },
  });
  if (compiled.isErr()) return "refused";
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  expect(editingView(applied.value.document)).toEqual(
    joinedView(document, blockId, nextBlockId),
  );
  const undone = applyDocumentOps(
    applied.value.document,
    applied.value.inverse,
  );
  if (undone.isErr()) throw undone.error;
  assertExactModel(undone.value.document, document);
  return "joined";
};

const markPending = (
  paragraph: Paragraph,
  kind: PendingMark,
  id: number,
): void => {
  paragraph.pPrMark = {
    kind,
    info: { id, author: "Earlier", date: "2026-01-02T03:04:05Z" },
  };
};

type Draw = {
  document: Document;
  pair: number;
  first: PendingMark | "asDrawn";
  next: PendingMark | "asDrawn";
};

type Tally = Map<string, number>;

const count = (tally: Tally, key: string): void => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/**
 * Two adjacent paragraphs of a drawn body, either or both given a mark pending
 * deletion where a review may leave one: not on a paragraph that ends its
 * container or a section.
 */
const pendingJoins = (tally: Tally) =>
  fc.property(
    fc.record({
      document: fc.oneof(documentArbitrary, reviewDocumentArbitrary),
      pair: fc.nat(),
      first: fc.constantFrom("asDrawn", ...PENDING_MARKS),
      next: fc.constantFrom("asDrawn", ...PENDING_MARKS),
    }),
    ({ document: drawn, pair, first, next }: Draw) => {
      const document = structuredClone(drawn);
      const paragraphs = storyParagraphs(document.package.document);
      const followed = (index: number) => {
        const location = paragraphs[index];
        const after = paragraphs[index + 1];
        return (
          location !== undefined &&
          after !== undefined &&
          sameBlockList(location.list, after.list) &&
          after.index === location.index + 1
        );
      };
      const pairs = paragraphs.flatMap((location, index) => {
        const after = paragraphs[index + 1];
        return followed(index) &&
          after !== undefined &&
          location.paragraph.paraId !== undefined &&
          after.paragraph.paraId !== undefined
          ? [{ index, leading: location.paragraph, trailing: after.paragraph }]
          : [];
      });
      fc.pre(pairs.length > 0);
      const chosen = pairs[pair % pairs.length];
      if (chosen === undefined) return;
      const { index, leading, trailing } = chosen;
      if (first !== "asDrawn" && leading.sectionProperties === undefined)
        markPending(leading, first, ADDED_REVISION_IDS);
      if (
        next !== "asDrawn" &&
        trailing.sectionProperties === undefined &&
        followed(index + 1)
      )
        markPending(trailing, next, ADDED_REVISION_IDS + 1);
      const pending = (paragraph: Paragraph) =>
        paragraph.pPrMark?.kind === "del" ||
        paragraph.pPrMark?.kind === "moveFrom";
      let shape = "live";
      if (pending(trailing)) shape = "pendingNext";
      if (pending(leading)) shape = "pendingFirst";
      for (const survivor of SURVIVORS) {
        const outcome = joinLaws(
          document,
          leading.paraId ?? "",
          trailing.paraId ?? "",
          survivor,
        );
        count(tally, `${shape}:${survivor}:${outcome}`);
      }
    },
  );

const paragraph = (
  paraId: string,
  text: string,
  formatting?: ParagraphFormatting,
): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
  ...(formatting === undefined ? {} : { formatting }),
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const inCell = (content: BlockContent[]): BlockContent[] => [
  {
    type: "table",
    rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
  },
  paragraph("0000000F", ""),
];

describe("a direct join leaves the editing view the editor shows", () => {
  describe.each(PENDING_MARKS)("after a mark pending as %s", (kind) => {
    const pendingFirst = () => {
      const first = paragraph("0000000A", "Alpha", {
        alignment: "start",
        styleId: "Heading1",
      });
      markPending(first, kind, 1);
      return first;
    };
    type Shape = {
      build: () => Document;
      view: ReturnType<typeof editingView>;
    };
    const shapes: Record<string, Shape> = {
      // The second paragraph ends the body and its section, empty.
      "an empty paragraph ending the body": {
        build: () =>
          documentOf([
            pendingFirst(),
            {
              ...paragraph("0000000B", "", {
                alignment: "center",
                keepNext: true,
              }),
              sectionProperties: { pageWidth: 12240 },
            },
          ]),
        view: [
          {
            text: "Alpha",
            properties: { alignment: "center", keepNext: true },
          },
        ],
      },
      "a paragraph with words": {
        build: () =>
          documentOf([
            pendingFirst(),
            paragraph("0000000B", "Beta", { alignment: "center" }),
            paragraph("0000000C", "Gamma"),
          ]),
        view: [
          { text: "AlphaBeta", properties: { alignment: "center" } },
          { text: "Gamma", properties: {} },
        ],
      },
      "a paragraph before a table": {
        build: () =>
          documentOf([
            pendingFirst(),
            paragraph("0000000B", "Beta", { alignment: "center" }),
            ...inCell([paragraph("0000000D", "Cell")]),
          ]),
        view: [
          { text: "AlphaBeta", properties: { alignment: "center" } },
          { text: "Cell", properties: {} },
          { text: "", properties: {} },
        ],
      },
      "a paragraph in a table cell": {
        build: () =>
          documentOf(
            inCell([
              pendingFirst(),
              paragraph("0000000B", "Beta", { alignment: "center" }),
            ]),
          ),
        view: [
          { text: "AlphaBeta", properties: { alignment: "center" } },
          { text: "", properties: {} },
        ],
      },
    };
    for (const [shape, { build, view }] of Object.entries(shapes)) {
      test.each(SURVIVORS)(
        `joined to ${shape}, keeping the %s, the view is unchanged`,
        (survivor) => {
          const document = build();
          // The two are one paragraph in the view: both paragraphs' words
          // and the second's properties, which its mark gives the pair.
          expect(editingView(document)).toEqual(view);
          expect(joinLaws(document, "0000000A", "0000000B", survivor)).toBe(
            "joined",
          );
        },
      );
    }

    test.each(SURVIVORS)(
      "a paragraph joined to one whose mark is pending, keeping the %s, shows its own properties",
      (survivor) => {
        const next = paragraph("0000000B", "Beta", { alignment: "end" });
        markPending(next, kind, 2);
        const document = documentOf([
          paragraph("0000000A", "Alpha", { alignment: "start" }),
          next,
          paragraph("0000000C", "Gamma", { alignment: "center" }),
        ]);
        expect(joinLaws(document, "0000000A", "0000000B", survivor)).toBe(
          "joined",
        );
        expect(joinedView(document, "0000000A", "0000000B")).toEqual([
          { text: "AlphaBetaGamma", properties: { alignment: "start" } },
        ]);
      },
    );
  });

  const GENERATED_RUNS = 200;
  const COVERAGE_RUNS = 200;

  test(
    "generated joins beside pending marks leave the view the editor shows",
    () => {
      assertProperty(pendingJoins(new Map()), {
        id: "generated joins beside pending marks leave the view the editor shows",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "generated joins reach every pending shape with either survivor",
    () => {
      const tally: Tally = new Map();
      checkCoverageSample(pendingJoins(tally), COVERAGE_RUNS);
      for (const shape of ["pendingFirst", "pendingNext", "live"])
        for (const survivor of SURVIVORS)
          expect(tally.get(`${shape}:${survivor}:joined`) ?? 0).toBeGreaterThan(
            0,
          );
    },
    propertyTestTimeout(120_000),
  );
});
