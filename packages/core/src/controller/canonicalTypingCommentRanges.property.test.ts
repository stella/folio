import { afterEach, expect, setSystemTime, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  getCommentAnchorsFromDoc,
  getTrackedChangeStatsFromDoc,
} from "../ai-edits/read";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import {
  type CommentMarkerSlot,
  visitCommentMarkers,
} from "../docx/commentAnchorIndex";
import { parseDocx } from "../docx/parser";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Comment, Document, ParagraphContent } from "../types/document";
import { readModel } from "../types/documentModel";

afterEach(() => setSystemTime());

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
type SourceOptions = {
  references: number;
  selection: "coveredText" | "wholeParagraph";
};
const sourceDocument = ({ references, selection }: SourceOptions) => {
  const atoms = Array.from(
    { length: references },
    (_, id) => ({ type: "commentReference", id }) satisfies ParagraphContent,
  );
  return {
    package: {
      document: {
        comments: Array.from(
          { length: Math.max(1, references) },
          (_, id) =>
            ({
              id,
              anchorKind: "range",
              author: "Reviewer",
              content: [
                {
                  type: "paragraph",
                  paraId: (0x20000000 + id).toString(16),
                  content: [run("Comment")],
                },
              ],
            }) satisfies Comment,
        ),
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            content: [
              ...(selection === "coveredText"
                ? [run("Prefix"), ...atoms.slice(1), run(" ")]
                : []),
              { type: "commentRangeStart", id: 0 },
              run("Commented"),
              { type: "commentRangeEnd", id: 0 },
              ...(selection === "coveredText" ? atoms.slice(0, 1) : atoms),
              run(" tail"),
            ],
          },
        ],
      },
    },
  } satisfies Document;
};

const savedView = async (buffer: ArrayBuffer) => {
  const document = await parseDocx(buffer, {
    preloadFonts: false,
    detectVariables: false,
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer.slice(0));
  const markers: CommentMarkerSlot["item"][] = [];
  visitCommentMarkers(document.package.document.content, ({ item }) =>
    markers.push(item),
  );
  return {
    definitions: document.package.document.comments,
    comments: reviewer.getComments(),
    anchors: getCommentAnchorsFromDoc(toProseDoc(document)),
    markers,
    text: toProseDoc(document).textContent,
  };
};

// Typing over a comment's whole words leaves the typed words outside its range.
test(
  "successive typing follows the replacement caret across comment boundaries",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("x", "é", "👍🏽"),
        fc.integer({ min: 0, max: 3 }),
        async (first, references) => {
          const acceptedViews = new Map<
            string,
            Awaited<ReturnType<typeof savedView>>
          >();
          for (const type of ["editing", "suggesting"] as const) {
            const source = sourceDocument({
              references,
              selection: "coveredText",
            });
            const session = createCanonicalSession(source).unwrap();
            session.setMode(
              type === "editing" ? { type } : { type, author: "Reviewer" },
            );
            let state = EditorState.create({ doc: session.projection.doc });
            let covered: { from: number; to: number } | undefined;
            state.doc.descendants((node, position) => {
              if (node.isText && node.text === "Commented")
                covered = { from: position, to: position + node.nodeSize };
            });
            if (covered === undefined)
              throw new Error("Missing covered source text");
            state = state.apply(
              state.tr.setSelection(
                TextSelection.create(state.doc, covered.from, covered.to),
              ),
            );
            const initial = await savedView(
              (
                await serializeCanonicalSave({
                  snapshot: session.captureSaveSnapshot(),
                })
              ).buffer,
            );
            const typed = [first, "y", "z"];
            for (const text of typed) {
              const { from, to } = state.selection;
              const prepared = session.prepareReplace(state, {
                from,
                to,
                text,
                semantic: "typing",
              });
              expect(
                prepared.isErr() ? prepared.error : undefined,
              ).toBeUndefined();
              const published = publishCanonicalProjection({
                session,
                state,
                commit: prepared.unwrap(),
              });
              expect(
                published.isErr() ? published.error : undefined,
              ).toBeUndefined();
              state = published.unwrap().state;
              expect(state.selection.empty).toBe(true);
            }
            const { ids: revisionIds } = getTrackedChangeStatsFromDoc(
              toProseDoc(session.document),
            );
            if (revisionIds.length > 0) {
              const rejected = createCanonicalSession(
                readModel(session.document),
              ).unwrap();
              const rejectedState = EditorState.create({
                doc: rejected.projection.doc,
              });
              publishCanonicalProjection({
                session: rejected,
                state: rejectedState,
                commit: rejected
                  .prepareResolve(rejectedState, {
                    revisionIds,
                    resolution: "reject",
                  })
                  .unwrap(),
              }).unwrap();
              expect(
                await savedView(
                  (
                    await serializeCanonicalSave({
                      snapshot: rejected.captureSaveSnapshot(),
                    })
                  ).buffer,
                ),
              ).toEqual(initial);
              const prepared = session.prepareResolve(state, {
                revisionIds,
                resolution: "accept",
              });
              expect(
                prepared.isErr() ? prepared.error : undefined,
              ).toBeUndefined();
              state = publishCanonicalProjection({
                session,
                state,
                commit: prepared.unwrap(),
              }).unwrap().state;
            }
            expect(state.doc.textContent).toBe(`Prefix ${typed.join("")} tail`);
            const accepted = await savedView(
              (
                await serializeCanonicalSave({
                  snapshot: session.captureSaveSnapshot(),
                })
              ).buffer,
            );
            // The typed words go outside the comment whose words they
            // replace, so accepting leaves that comment nothing to cover: it
            // goes, and every other comment stays as it was.
            const others = <T extends { id: number }>(items: readonly T[]) =>
              items.filter(({ id }) => id !== 0);
            expect(accepted.definitions ?? []).toEqual(
              others(initial.definitions ?? []),
            );
            expect(accepted.markers).toEqual(others(initial.markers));
            expect(
              accepted.comments.map(({ id, text, anchoredText }) => ({
                id,
                text,
                anchoredText,
              })),
            ).toEqual(
              others(initial.comments).map(({ id, text, anchoredText }) => ({
                id,
                text,
                anchoredText,
              })),
            );
            acceptedViews.set(type, accepted);
          }
          expect([...acceptedViews.keys()]).toEqual(["editing", "suggesting"]);
          expect(acceptedViews.get("editing")).toEqual(
            acceptedViews.get("suggesting"),
          );
        },
      ),
      {
        seed: 1404001,
        numRuns: 12,
        id: "typing after replacement preserves caret order across comment anchors",
      },
    );
  },
  propertyTestTimeout(30_000),
);

// Typed over directly, the ranged comment goes with its words, definition and
// markers alike. Tracked, the markers are deleted with the words: accepted,
// the comment has nothing left to cover and goes as it does directly;
// rejected, everything returns. A reference whose comment covers nothing has
// no words to go with: it moves to the start of the typed text, directly and
// once accepted.
for (const references of [0, 1, 2, 3])
  for (const type of ["editing", "suggesting"] as const) {
    test(`${type}, ${references} references: replacing the whole mixed-coverage paragraph takes its comment and keeps the bare references`, async () => {
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const source = sourceDocument({
        references,
        selection: "wholeParagraph",
      });
      const session = createCanonicalSession(source).unwrap();
      session.setMode(
        type === "editing" ? { type } : { type, author: "Reviewer" },
      );
      let state = EditorState.create({ doc: session.projection.doc });
      state = state.apply(
        state.tr.setSelection(
          TextSelection.create(state.doc, 1, state.doc.content.size - 1),
        ),
      );
      const before = session.document;
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
      });
      const original = await savedView(saved.buffer);
      const result = session.prepareReplace(state, {
        from: state.selection.from,
        to: state.selection.to,
        text: "x",
        semantic: "typing",
      });
      state = publishCanonicalProjection({
        session,
        state,
        commit: result.unwrap(),
      }).unwrap().state;
      if (type === "suggesting") {
        const resolve = (resolution: "accept" | "reject") => {
          state = publishCanonicalProjection({
            session,
            state,
            commit: session
              .prepareResolve(state, {
                revisionIds: [...getTrackedChangeStatsFromDoc(state.doc).ids],
                resolution,
              })
              .unwrap(),
          }).unwrap().state;
        };
        resolve("reject");
        expect(
          await savedView(
            (
              await serializeCanonicalSave({
                snapshot: session.captureSaveSnapshot(),
              })
            ).buffer,
          ),
        ).toEqual(original);
        state = publishCanonicalProjection({
          session,
          state,
          commit: session.prepareUndo(state).unwrap(),
        }).unwrap().state;
        resolve("accept");
      }
      const typed = await savedView(
        (
          await serializeCanonicalSave({
            snapshot: session.captureSaveSnapshot(),
          })
        ).buffer,
      );
      // The ranged comment 0 goes; each bare reference stays, now before
      // the typed text.
      const kept = ({ id }: { id: number }) => id !== 0;
      expect(typed.definitions ?? []).toEqual(
        (original.definitions ?? []).filter(kept),
      );
      expect(typed.markers).toEqual(original.markers.filter(kept));
      expect(typed.comments.map(({ id }) => id)).toEqual(
        original.comments.map(({ id }) => id).filter((id) => id !== 0),
      );
      expect(typed.text).toBe("x");
      while (session.canUndo)
        state = publishCanonicalProjection({
          session,
          state,
          commit: session.prepareUndo(state).unwrap(),
        }).unwrap().state;
      expect(session.document).toEqual(before);
    });
  }
