import { expect, test, setDefaultTimeout } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  paragraphVisibleText,
  physicalOffsetAtVisibleOffset,
  storyBody,
  type OpStory,
} from "@stll/docx-core/ops";

import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { canonicalReviewBlocks, storyRevisionIds } from "../../../../test/reviewProjection";
import { reviewDifferences } from "../../../../test/reviewDifferences";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import { createCanonicalSession, type CanonicalCommit } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(240_000));

const STORIES = [
  { kind: "header", rId: "rIdHeader" },
  { kind: "footer", rId: "rIdFooterDefault" },
  { kind: "footer", rId: "rIdFooterFirst" },
  { kind: "footer", rId: "rIdFooterEven" },
  { kind: "footnote", id: 12 },
  { kind: "endnote", id: 13 },
] as const satisfies readonly OpStory[];

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [
    { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "" }] },
    { type: "run", formatting: { bold: true }, content: [{ type: "text", text }] },
  ],
});

const fixture = (): Document => ({
  package: {
    document: {
      content: [
        {
          ...paragraph("12345678", "Body"),
          content: [
            {
              type: "run",
              content: [
                { type: "text", text: "Body" },
                { type: "footnoteRef", id: 12 },
                { type: "endnoteRef", id: 13 },
              ],
            },
          ],
          sectionProperties: {
            titlePg: true,
            headerReferences: [{ type: "default", rId: "rIdHeader" }],
            footerReferences: [
              { type: "default", rId: "rIdFooterDefault" },
              { type: "first", rId: "rIdFooterFirst" },
              { type: "even", rId: "rIdFooterEven" },
            ],
          },
        },
      ],
    },
    headers: new Map([
      [
        "rIdHeader",
        { type: "header", hdrFtrType: "default", content: [paragraph("23456789", "A😀éB")] },
      ],
    ]),
    footers: new Map([
      [
        "rIdFooterDefault",
        { type: "footer", hdrFtrType: "default", content: [paragraph("34567890", "A😀éB")] },
      ],
      [
        "rIdFooterFirst",
        { type: "footer", hdrFtrType: "first", content: [paragraph("45678901", "A😀éB")] },
      ],
      [
        "rIdFooterEven",
        { type: "footer", hdrFtrType: "even", content: [paragraph("56789012", "A😀éB")] },
      ],
    ]),
    footnotes: [{ type: "footnote", id: 12, content: [paragraph("67890123", "A😀éB")] }],
    endnotes: [{ type: "endnote", id: 13, content: [paragraph("78901234", "A😀éB")] }],
  },
});

const asMainStory = (document: Document, story: OpStory): Document => ({
  package: { document: { content: storyBody(document, story).content } },
});

const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(next.doc.eq(commit.projection.doc)).toBe(true);
  expect(commit.publish().isOk()).toBe(true);
  return next;
};

const assertSavedStories = async (document: Document) => {
  const reopened = await parseDocx(await createDocx(document), { preloadFonts: false });
  for (const story of [OP_STORIES.MAIN, ...STORIES]) {
    expect(reviewDifferences(asMainStory(document, story), asMainStory(reopened, story))).toEqual({
      messages: [],
      omitted: 0,
    });
  }
  return reopened;
};

test("generated tracked secondary replacements preserve shared history and saved review outcomes", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(
        fc.record({
          index: fc.nat(50),
          count: fc.nat(4),
          text: fc.constantFrom("", "x", "😀", "é"),
          author: fc.constantFrom("Story author", "Other author"),
        }),
        { minLength: 4, maxLength: 9 },
      ),
      async (trace) => {
        for (const story of STORIES) {
          const direct = createCanonicalSession(fixture()).unwrap();
          const tracked = createCanonicalSession(fixture()).unwrap();
          tracked.setMode({ type: "suggesting", author: "Story author" });
          let directState = EditorState.create({
            schema,
            doc: direct.projectStory(story).unwrap().doc,
          });
          let trackedState = EditorState.create({
            schema,
            doc: tracked.projectStory(story).unwrap().doc,
          });
          const baseline = tracked.document;
          const snapshots = [baseline];
          const selections = [trackedState.selection.toJSON()];
          for (const [step, input] of trace.entries()) {
            if (step === 2) {
              // Body and story edits share one journal even when undo is requested by the story editor.
              tracked.setMode({ type: "editing" });
              for (const session of [direct, tracked]) {
                const bodyState = EditorState.create({ schema, doc: session.projection.doc });
                accept(
                  bodyState,
                  session.prepareReplace(bodyState, { from: 1, to: 1, text: "Journal" }).unwrap(),
                );
              }
              snapshots.push(tracked.document);
              selections.push(trackedState.selection.toJSON());
            }
            tracked.setMode({ type: "suggesting", author: input.author });
            const before = tracked.document;
            const current = storyBody(tracked.document, story).content.at(0);
            if (current?.type !== "paragraph") panic("The story generator lost its paragraph.");
            const gaps = [0];
            let offset = 0;
            for (const character of paragraphVisibleText(current)) {
              offset += character.length;
              gaps.push(offset);
            }
            const index = input.index % gaps.length;
            const from = gaps.at(index) ?? panic("The story generator lost its start gap.");
            const to =
              gaps.at(Math.min(index + input.count, gaps.length - 1)) ??
              panic("The story generator lost its end gap.");
            if (from === to && input.text === "") continue;
            directState = accept(
              directState,
              direct
                .prepareReplace(directState, {
                  from: direct
                    .projectStory(story)
                    .unwrap()
                    .positionAt({
                      story,
                      blockId: current.paraId ?? panic("Story paragraph id missing."),
                      offset: from,
                    })
                    .unwrap(),
                  to: direct
                    .projectStory(story)
                    .unwrap()
                    .positionAt({
                      story,
                      blockId: current.paraId ?? panic("Story paragraph id missing."),
                      offset: to,
                    })
                    .unwrap(),
                  text: input.text,
                  story,
                })
                .unwrap(),
            );
            trackedState = accept(
              trackedState,
              tracked
                .prepareReplace(trackedState, {
                  from: tracked
                    .projectStory(story)
                    .unwrap()
                    .positionAt({
                      story,
                      blockId: current.paraId ?? panic("Story paragraph id missing."),
                      offset: physicalOffsetAtVisibleOffset(current, from),
                    })
                    .unwrap(),
                  to: tracked
                    .projectStory(story)
                    .unwrap()
                    .positionAt({
                      story,
                      blockId: current.paraId ?? panic("Story paragraph id missing."),
                      offset: physicalOffsetAtVisibleOffset(current, to),
                    })
                    .unwrap(),
                  text: input.text,
                  story,
                })
                .unwrap(),
            );
            snapshots.push(tracked.document);
            selections.push(trackedState.selection.toJSON());
            for (const other of [OP_STORIES.MAIN, ...STORIES].filter(
              (candidate) => candidate !== story,
            )) {
              assertExactModel(
                storyBody(tracked.document, other).content,
                storyBody(before, other).content,
              );
            }
          }
          const pending = tracked.document;
          const reopenedPending = await assertSavedStories(pending);
          const revisionIds = storyRevisionIds(asMainStory(pending, story));
          const pendingParagraph = storyBody(pending, story).content.at(0);
          if (pendingParagraph?.type !== "paragraph")
            panic("The authored story lost its paragraph.");
          for (const content of pendingParagraph.content) {
            if (content.type === "insertion" || content.type === "deletion")
              expect(["Story author", "Other author"].includes(content.info.author)).toBe(true);
          }
          for (const decision of ["accept", "reject"] as const) {
            if (revisionIds.length === 0) continue;
            const resolved = applyDocumentOps(pending, [
              { type: DOCUMENT_OP_TYPES.RESOLVE_REVISION, story, revisionIds, decision },
            ]).unwrap();
            const expected = decision === "accept" ? direct.document : baseline;
            expect(canonicalReviewBlocks(storyBody(resolved.document, story).content)).toEqual(
              canonicalReviewBlocks(storyBody(expected, story).content),
            );
            const reopenedResolved = applyDocumentOps(reopenedPending, [
              { type: DOCUMENT_OP_TYPES.RESOLVE_REVISION, story, revisionIds, decision },
            ]).unwrap();
            expect(
              reviewDifferences(
                asMainStory(expected, story),
                asMainStory(reopenedResolved.document, story),
              ),
            ).toEqual({ messages: [], omitted: 0 });
            for (const other of [OP_STORIES.MAIN, ...STORIES].filter(
              (candidate) => candidate !== story,
            )) {
              assertExactModel(
                storyBody(resolved.document, other).content,
                storyBody(pending, other).content,
              );
            }
            const undone = applyDocumentOps(resolved.document, resolved.inverse).unwrap();
            assertExactModel(undone.document, pending);
            assertExactModel(
              applyDocumentOps(undone.document, undone.inverse).unwrap().document,
              resolved.document,
            );
            await assertSavedStories(resolved.document);
          }
          for (let index = snapshots.length - 2; index >= 0; index--) {
            trackedState = accept(trackedState, tracked.prepareUndo(trackedState, story).unwrap());
            assertExactModel(tracked.document, snapshots.at(index));
            expect(trackedState.selection.toJSON()).toEqual(selections.at(index));
          }
          for (let index = 1; index < snapshots.length; index++) {
            trackedState = accept(trackedState, tracked.prepareRedo(trackedState, story).unwrap());
            assertExactModel(tracked.document, snapshots.at(index));
            expect(trackedState.selection.toJSON()).toEqual(selections.at(index));
          }
        }
      },
    ),
    {
      numRuns: 12,
      id: "generated tracked secondary replacements preserve shared history and saved review outcomes",
    },
  );
});
