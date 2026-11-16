/** Automatic note marks survive resolution without restoring removed revision payload. */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  Paragraph,
  Run,
  TrackedRunContent,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { IDENTITY_SPACES, identityKeysIn, slotKey } from "../ids";
import { leafSpans } from "../leaves";
import { automaticNoteMarkRuns } from "../noteMarks";
import { isTrackedWrapper, type TrackedWrapper } from "../review";
import { storyBody } from "../stories";
import { DOCUMENT_OP_TYPES, type DocumentOp, type NoteStory } from "../types";

setDefaultTimeout(propertyTestTimeout(5_000));

const DATE = "2026-01-02T03:04:05Z";
const NOTE_ID = "10000011";
const FOLLOWER_ID = "10000012";

const paragraph = (
  paraId: string,
  content: TrackedRunContent[],
): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

const documentOf = (
  story: NoteStory,
  content: TrackedRunContent[],
): Document => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph("10000001", [
            {
              type: "run",
              content: [
                {
                  type:
                    story.kind === "footnote" ? "footnoteRef" : "endnoteRef",
                  id: story.id,
                },
              ],
            },
          ]),
        ],
      },
    },
  };
  const blocks = [
    paragraph(NOTE_ID, content),
    paragraph(FOLLOWER_ID, [
      { type: "run", content: [{ type: "text", text: "Follower." }] },
    ]),
  ];
  if (story.kind === "footnote")
    document.package.footnotes = [
      { type: "footnote", id: story.id, content: blocks },
    ];
  else
    document.package.endnotes = [
      { type: "endnote", id: story.id, content: blocks },
    ];
  return document;
};

const applied = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops).unwrap();
  expect(contractViolation(result.document)).toBeUndefined();
  expect(
    applyDocumentOps(result.document, result.inverse).unwrap().document,
  ).toEqual(document);
  return result;
};

const markerRun = (kind: NoteStory["kind"]): Run => ({
  type: "run",
  formatting: {
    styleId: kind === "footnote" ? "FootnoteReference" : "EndnoteReference",
    bold: true,
    color: { rgb: "336699" },
  },
  content: [
    { type: "noteMarker", kind },
    { type: "text", text: "Removed words." },
  ],
});

const wrapped = (
  types: readonly TrackedWrapper["type"][],
  run: Run,
): TrackedRunContent[] => {
  let content: TrackedRunContent[] = [run];
  for (const [reverseIndex, type] of types.toReversed().entries())
    content = [
      {
        type,
        info: {
          id: types.length - reverseIndex + 9,
          author: "Reviewer",
          date: DATE,
        },
        content,
      },
    ];
  return content;
};

const kinds = ["footnote", "endnote"] as const;
const decisions = ["accept", "reject"] as const;
const wrapperTypes = ["insertion", "deletion", "moveFrom", "moveTo"] as const;

for (const kind of kinds) {
  test(`${kind}: accepting a deleted first paragraph keeps only its source automatic mark and undoes exactly`, () => {
    const story = { kind, id: 2 };
    const source = markerRun(kind);
    const document = documentOf(story, [
      {
        type: "deletion",
        info: { id: 406, author: "Reviewer", date: DATE },
        content: [
          { ...source, content: [{ type: "noteMarker", kind }] },
          { type: "run", content: [{ type: "text", text: "Removed words." }] },
        ],
      },
    ]);
    const first = storyBody(document, story).content.at(0);
    if (first?.type !== "paragraph")
      throw new Error("Missing first note paragraph");
    first.formatting = { alignment: "start", keepNext: true };
    first.pPrMark = {
      kind: "del",
      info: { id: 405, author: "Reviewer", date: DATE },
    };
    const before = structuredClone(document);
    const result = applied(document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds: [405, 406],
        decision: "accept",
      },
    ]);
    const note = storyBody(result.document, story);
    expect(note.content).toEqual([
      paragraph(FOLLOWER_ID, [
        { ...source, content: [{ type: "noteMarker", kind }] },
        { type: "run", content: [{ type: "text", text: "Follower." }] },
      ]),
    ]);
    expect(document).toEqual(before);
  });

  test(`${kind}: a raw inline replacement retains only the automatic marker run`, () => {
    const story = { kind, id: 2 };
    const source = markerRun(kind);
    const content = wrapped(["insertion", "deletion"], source);
    const document = documentOf(story, content);
    const result = applied(document, [
      {
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story,
        blockId: NOTE_ID,
        expected: content,
        content: [
          { type: "run", content: [{ type: "text", text: "Replacement." }] },
        ],
      },
    ]);
    expect(storyBody(result.document, story).content.at(0)).toEqual(
      paragraph(NOTE_ID, [
        { ...source, content: [{ type: "noteMarker", kind }] },
        { type: "run", content: [{ type: "text", text: "Replacement." }] },
      ]),
    );
  });
}

for (const kind of kinds)
  for (const decision of decisions)
    test(`${kind} ${decision}: nested revision resolution keeps exactly one formatted automatic mark and only pending wrappers`, async () => {
      await assertProperty(
        fc.property(
          fc.array(fc.constantFrom(...wrapperTypes), {
            minLength: 1,
            maxLength: 4,
          }),
          fc.array(fc.boolean(), { minLength: 4, maxLength: 4 }),
          fc.boolean(),
          fc.boolean(),
          (types, selected, bold, resolveFormatting) => {
            const story = { kind, id: 2 };
            const formatting = { ...markerRun(kind).formatting, bold };
            const previousFormatting = { ...formatting, bold: !bold };
            const source = {
              ...markerRun(kind),
              formatting,
              propertyChanges: [
                {
                  type: "runPropertyChange",
                  info: { id: 99, author: "Formatter", date: DATE },
                  previousFormatting,
                  currentFormatting: formatting,
                },
              ],
            } satisfies Run;
            const document = documentOf(story, wrapped(types, source));
            const before = structuredClone(document);
            const revisionIds = [
              ...(resolveFormatting ? [99] : []),
              ...types.flatMap((_, index) =>
                index === 0 || selected[index] ? [index + 10] : [],
              ),
            ];
            const resolves = (id: number) => revisionIds.includes(id);
            const removesWords = types.some(
              (type, index) =>
                resolves(index + 10) &&
                (type === "insertion" || type === "moveTo") ===
                  (decision === "reject"),
            );
            const result = applied(document, [
              {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story,
                revisionIds,
                decision,
              },
            ]);
            const note = storyBody(result.document, story);
            const first = note.content.at(0);
            if (first?.type !== "paragraph")
              throw new Error("Missing resolved note paragraph");
            const spans = leafSpans(first.content);
            expect(
              spans.filter(({ node }) => node.type === "noteMarker"),
            ).toHaveLength(1);
            expect(automaticNoteMarkRuns(first.content)).toEqual([
              {
                type: "run",
                formatting:
                  resolveFormatting && decision === "reject"
                    ? previousFormatting
                    : formatting,
                content: [{ type: "noteMarker", kind }],
                ...(!resolveFormatting && {
                  propertyChanges: source.propertyChanges,
                }),
              },
            ]);
            const words = spans
              .flatMap(({ node }) => (node.type === "text" ? [node.text] : []))
              .join("");
            expect(words).toBe(removesWords ? "" : "Removed words.");
            const marker = spans.find(({ node }) => node.type === "noteMarker");
            if (!marker)
              throw new Error("Resolution removed the automatic mark");
            expect(
              marker.ancestors
                .filter(isTrackedWrapper)
                .map(({ info }) => info.id),
            ).toEqual(
              types.flatMap((_, index) =>
                resolves(index + 10) ? [] : [index + 10],
              ),
            );
            const keys = new Set(identityKeysIn(first.content));
            for (const id of revisionIds)
              expect(
                keys.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
              ).toBe(false);
            const allIds = [99, ...types.map((_, index) => index + 10)];
            const sequential = applied(result.document, [
              {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story,
                revisionIds: allIds.filter((id) => !resolves(id)),
                decision,
              },
            ]);
            const together = applied(document, [
              {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story,
                revisionIds: allIds,
                decision,
              },
            ]);
            expect(sequential.document).toEqual(together.document);
            expect(document).toEqual(before);
          },
        ),
        {
          id: "automatic note mark resolution preserves source ownership",
          numRuns: 50,
        },
      );
    });

for (const kind of kinds) {
  test(`${kind}: a pending insertion keeps its produced seam metadata when an enclosing deletion is accepted`, () => {
    const story = { kind, id: 2 };
    const sourceMark = {
      ...markerRun(kind),
      content: [{ type: "noteMarker", kind }],
    } satisfies Run;
    const document = documentOf(story, [
      sourceMark,
      { type: "run", content: [{ type: "text", text: "AB" }] },
    ]);
    const inserted = applied(document, [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story, blockId: NOTE_ID, offset: 1 },
        text: "Inserted words.",
        runProps: "inherit",
        revision: { id: 10, author: "Writer", date: DATE },
      },
    ]);
    const first = storyBody(inserted.document, story).content.at(0);
    if (first?.type !== "paragraph")
      throw new Error("Missing insertion paragraph");
    const pending = first.content.find((node) => node.type === "insertion");
    if (pending?.type !== "insertion")
      throw new Error("Insertion producer returned no revision");
    expect(pending.resolutionJoins?.remove).toBeGreaterThan(0);
    // The imported source mark stands inside the producer's pending seam.
    const pendingWithMark = {
      ...pending,
      content: [sourceMark, ...pending.content],
    };
    const enclosed = applied(inserted.document, [
      {
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story,
        blockId: NOTE_ID,
        expected: first.content,
        content: [
          {
            type: "deletion",
            info: { id: 20, author: "Reviewer", date: DATE },
            content: [pendingWithMark],
          },
        ],
      },
    ]);
    const partial = applied(enclosed.document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds: [20],
        decision: "accept",
      },
    ]);
    const after = storyBody(partial.document, story).content.at(0);
    if (after?.type !== "paragraph")
      throw new Error("Missing resolved paragraph");
    const retained = after.content.find((node) => node.type === "insertion");
    expect(retained).toEqual({
      ...pendingWithMark,
      content: automaticNoteMarkRuns(pendingWithMark.content),
    });
    for (const decision of decisions) {
      const sequential = applied(partial.document, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story,
          revisionIds: [10],
          decision,
        },
      ]);
      const together = applied(enclosed.document, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story,
          revisionIds: [10],
          decision,
        },
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story,
          revisionIds: [20],
          decision: "accept",
        },
      ]);
      expect(sequential.document).toEqual(together.document);
    }
  });
}
