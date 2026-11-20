import { panic } from "better-result";
import { afterEach, describe, expect, setSystemTime, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  getCommentAnchorsFromDoc,
  getTrackedChangesFromDoc,
} from "../ai-edits/read";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import {
  canonicalCommentBody,
  compileCanonicalComments,
} from "../canonical/canonicalComments";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import {
  type CommentMarkerSlot,
  visitCommentMarkers,
} from "../docx/commentAnchorIndex";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document, InlineSdt, Run } from "../types/document";
import { readModel } from "../types/documentModel";

test("compiler-created boundary points survive native replacement or refuse unchanged", async () => {
  setSystemTime(new Date("2026-01-02T03:04:05Z"));
  const sourceText = "alpha beta";
  const markerFacts = (document: Document) => {
    const markers: CommentMarkerSlot["item"][] = [];
    visitCommentMarkers(document.package.document.content, ({ item }) =>
      markers.push(item),
    );
    return markers;
  };
  for (const representation of ["point", "range"] as const)
    for (const offset of [0, sourceText.length])
      for (const mode of ["editing", "suggesting"] as const)
        for (const scope of ["whole", "boundary", "carrier"] as const) {
          const source = fromMarkdown(sourceText);
          const paragraph = source.package.document.content.at(0);
          if (paragraph?.type !== "paragraph")
            return panic("Produced paragraph is missing");
          paragraph.paraId = "12345678";
          const at = {
            story: "main" as const,
            blockId: paragraph.paraId,
            offset,
          };
          const created = compileCanonicalComments({
            document: source,
            command: {
              type: "create",
              comment: {
                id: 0,
                author: "Reviewer",
                content: canonicalCommentBody(source, "Point"),
              },
              anchor:
                representation === "point"
                  ? { kind: "point", at }
                  : { kind: "range", from: at, to: at },
            },
          }).unwrap();
          const session = createCanonicalSession(created.document).unwrap();
          session.setMode(
            mode === "editing"
              ? { type: "editing" }
              : { type: "suggesting", author: "Reviewer" },
          );
          let state = EditorState.create({ doc: session.projection.doc });
          const publish = (commit: CanonicalCommit) => {
            state = publishCanonicalProjection({
              state,
              commit,
              session,
            }).unwrap().state;
          };
          const block = session.projection.firstParagraph;
          const carriers: { from: number; to: number }[] = [];
          let textStart: number | undefined;
          state.doc.descendants((node, position) => {
            if (
              node.type.name === "rangeAnchor" ||
              node.type.name === "commentReference"
            )
              carriers.push({ from: position, to: position + node.nodeSize });
            if (node.isText && node.text === sourceText) textStart = position;
          });
          const carrier = carriers.at(0);
          if (
            block === undefined ||
            carrier === undefined ||
            textStart === undefined
          )
            return panic("Produced point source is missing");
          const range = (() => {
            switch (scope) {
              case "whole":
                return {
                  from: block.start,
                  to: block.start + block.node.content.size,
                };
              case "boundary":
                return offset === 0
                  ? {
                      from: block.start,
                      to: textStart + sourceText.indexOf(" "),
                    }
                  : {
                      from: textStart + sourceText.indexOf(" ") + 1,
                      to: block.start + block.node.content.size,
                    };
              case "carrier":
                return carrier;
              default: {
                const unhandled: never = scope;
                return panic(String(unhandled));
              }
            }
          })();
          const before = session.document;
          const anchors = getCommentAnchorsFromDoc(state.doc);
          const markers = markerFacts(created.document);
          const snapshot = session.captureSaveSnapshot();
          const saved = await serializeCanonicalSave({ snapshot });
          let committed = 0;
          session.onCommitted(() => committed++);
          expect(
            before.package.document.comments?.map(({ id, anchorKind }) => ({
              id,
              anchorKind,
            })),
          ).toEqual([{ id: 0, anchorKind: "point" }]);
          const prepared = session.prepareReplace(state, {
            ...range,
            text: "new",
          });
          if (prepared.isErr()) {
            expect(prepared.error.ground).toEqual({
              gap: "comment-replacement-alignment",
            });
            expect(session.document).toBe(before);
            expect(session.version).toBe(0);
            expect(session.canUndo).toBe(false);
            expect(committed).toBe(0);
            expect(session.captureSaveSnapshot()).toEqual(snapshot);
            expect(
              new Uint8Array(
                (
                  await serializeCanonicalSave({
                    snapshot: session.captureSaveSnapshot(),
                  })
                ).buffer,
              ),
            ).toEqual(new Uint8Array(saved.buffer));
            continue;
          }
          publish(prepared.value);
          if (mode === "suggesting")
            publish(
              session
                .prepareResolve(state, {
                  revisionIds: getTrackedChangesFromDoc(state.doc).map(
                    ({ id }) => id,
                  ),
                  resolution: "accept",
                })
                .unwrap(),
            );
          expect(session.document.package.document.comments).toEqual(
            before.package.document.comments,
          );
          expect(markerFacts(readModel(session.document))).toEqual(markers);
          expect(getCommentAnchorsFromDoc(state.doc)).toEqual(anchors);
          const reopened = createCanonicalSession(
            await parseDocx(
              (
                await serializeCanonicalSave({
                  snapshot: session.captureSaveSnapshot(),
                })
              ).buffer,
              { preloadFonts: false },
            ),
          ).unwrap();
          expect(
            reopened.document.package.document.comments?.map(
              ({ id, anchorKind }) => ({ id, anchorKind }),
            ),
          ).toEqual([{ id: 0, anchorKind: "point" }]);
          expect(markerFacts(readModel(reopened.document))).toEqual(markers);
          expect(getCommentAnchorsFromDoc(reopened.projection.doc)).toEqual(
            anchors,
          );
        }
});

test("DOCX-produced nested control comments are guarded directly and kept on deleted words when tracked", async () => {
  const text = (value: string) =>
    ({ type: "run", content: [{ type: "text", text: value }] }) satisfies Run;
  const control = (id: number, content: InlineSdt["content"]) =>
    ({
      type: "inlineSdt",
      properties: { sdtType: "richText", id },
      content,
    }) satisfies InlineSdt;
  for (const nesting of [1, 2]) {
    const inner = control(9, [
      text("Before "),
      { type: "commentRangeStart", id: 0 },
      text("time"),
      { type: "commentRangeEnd", id: 0 },
      { type: "commentReference", id: 0 },
      text(" after"),
    ]);
    const source = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "12345678",
              content: [
                text("Outside "),
                nesting === 1 ? inner : control(10, [inner]),
                text(" outside."),
              ],
            },
          ],
          comments: [
            {
              id: 0,
              author: "Reviewer",
              content: [
                {
                  type: "paragraph",
                  paraId: "23456789",
                  content: [text("Note")],
                },
              ],
            },
          ],
        },
      },
    } satisfies Document;
    // A real DOCX source keeps its markers inside the control. Creating a new
    // semantic comment splits that control and does not exercise this shape.
    const parsed = await parseDocx(await createDocx(source), {
      preloadFonts: false,
    });
    for (const mode of ["editing", "suggesting"] as const) {
      for (const placement of ["inside", "crossing", "covered"] as const) {
        const session = createCanonicalSession(parsed).unwrap();
        session.setMode(
          mode === "editing"
            ? { type: "editing" }
            : { type: "suggesting", author: "Reviewer" },
        );
        let state = EditorState.create({ doc: session.projection.doc });
        const publish = (commit: CanonicalCommit) => {
          state = publishCanonicalProjection({
            state,
            commit,
            session,
          }).unwrap().state;
        };
        const controls: { from: number; to: number }[] = [];
        let covered: { from: number; to: number } | undefined;
        state.doc.descendants((node, position) => {
          if (node.type.name === "sdt")
            controls.push({
              from: position + 1,
              to: position + node.nodeSize - 1,
            });
          if (node.isText && node.text === "time")
            covered = { from: position, to: position + node.nodeSize };
        });
        const inside = controls.at(-1);
        const paragraph = session.projection.firstParagraph;
        if (
          inside === undefined ||
          covered === undefined ||
          paragraph === undefined
        )
          return panic("Produced control source is missing");
        expect(
          getCommentAnchorsFromDoc(state.doc).map(({ quote }) => quote),
        ).toEqual(["time"]);
        const range = (() => {
          switch (placement) {
            case "inside":
              return { from: inside.from + 2, to: inside.to - 2 };
            case "crossing":
              return {
                from: paragraph.start,
                to: paragraph.start + paragraph.node.content.size,
              };
            case "covered":
              return covered;
            default: {
              const unhandled: never = placement;
              return panic(String(unhandled));
            }
          }
        })();
        const before = session.document;
        const snapshot = session.captureSaveSnapshot();
        let committed = 0;
        session.onCommitted(() => committed++);
        const result = session.prepareReplace(state, {
          ...range,
          text: "schedule",
        });
        if (placement === "inside" && mode === "suggesting") {
          // Tracked, the deleted words keep the comment: rejecting gives it
          // back where it was. A selection taking the comment's markers
          // still refuses.
          expect(result.isOk()).toBe(true);
          publish(result.unwrap());
          publish(
            session
              .prepareResolve(state, {
                revisionIds: getTrackedChangesFromDoc(state.doc).map(
                  ({ id }) => id,
                ),
                resolution: "reject",
              })
              .unwrap(),
          );
          expect(
            getCommentAnchorsFromDoc(state.doc).map(({ quote }) => quote),
          ).toEqual(["time"]);
          continue;
        }
        if (placement !== "covered") {
          expect(result.isErr()).toBe(true);
          if (result.isErr())
            expect(result.error.ground).toEqual({
              gap: "comment-replacement-alignment",
            });
          expect(session.document).toBe(before);
          expect(session.version).toBe(0);
          expect(session.canUndo).toBe(false);
          expect(committed).toBe(0);
          expect(session.captureSaveSnapshot()).toEqual(snapshot);
          const executor = new CanonicalPublicOperations({
            session,
            getState: () => state,
            publish: (commit) => {
              publish(commit);
              return true;
            },
          });
          const applied = executor.apply({
            snapshot: createFolioAIEditSnapshot(state.doc),
            batch: {
              version: 1,
              mode: "direct",
              operations: [
                {
                  id: "replace",
                  type: "replaceBlock",
                  blockId: "12345678",
                  text: "New agreement.",
                },
              ],
            },
          });
          expect(applied.applied).toEqual([]);
          expect(applied.skipped).toMatchObject([
            { canonicalRefusal: { gap: "comment-replacement-alignment" } },
          ]);
          expect(session.document).toBe(before);
          expect(session.version).toBe(0);
          expect(committed).toBe(0);
          continue;
        }
        expect(result.isOk()).toBe(true);
        publish(result.unwrap());
        if (mode === "suggesting")
          publish(
            session
              .prepareResolve(state, {
                revisionIds: getTrackedChangesFromDoc(state.doc).map(
                  ({ id }) => id,
                ),
                resolution: "accept",
              })
              .unwrap(),
          );
        // The new words go outside the range of the comment whose words
        // they replace, which leaves it nothing to cover.
        expect(
          getCommentAnchorsFromDoc(state.doc).map(({ quote }) => quote),
        ).toEqual([]);
        expect(state.doc.textContent).toContain("Before schedule after");
      }
    }
  }
});

const SOURCE = "The Supplier delivers the goods on time and in good order.";
const REPLACEMENT =
  "Invoice delivery payment agreement supplier buyer payment term.";
afterEach(() => setSystemTime());
const COMMENT_CASES = {
  prefix: { quotes: ["The"], type: "alignment" },
  interior: { quotes: ["time"], type: "alignment" },
  suffix: { quotes: ["order."], type: "supported" },
  whole: { quotes: [SOURCE], type: "supported" },
  disjoint: { quotes: ["The", "time"], type: "alignment" },
  nested: { quotes: [SOURCE, "time"], type: "alignment" },
  overlapping: { quotes: ["on time", "time and"], type: "alignment" },
} as const;

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    session,
    author: "Reviewer",
  });

const apply = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes",
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(1);
  return result;
};

const comments = (reviewer: FolioDocxReviewer) =>
  reviewer.getComments().map(({ id, text, anchoredText, replies }) => ({
    id,
    text,
    anchoredText,
    replies: replies.map(({ id: replyId, text: replyText }) => ({
      id: replyId,
      text: replyText,
    })),
  }));

describe("canonical replacement comments from the public producer", () => {
  for (const offset of [0, 4, SOURCE.length]) {
    test(`point at ${offset}: replacement retains its carriers or refuses unchanged`, async () => {
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const model = fromMarkdown(SOURCE);
      const paragraph = model.package.document.content.at(0);
      if (paragraph?.type !== "paragraph")
        throw new Error("Source paragraph is missing");
      paragraph.paraId = "12345678";
      const compiled = compileCanonicalComments({
        document: model,
        command: {
          type: "create",
          comment: {
            id: 0,
            author: "Reviewer",
            content: canonicalCommentBody(model, "Point"),
          },
          anchor: {
            kind: "point",
            at: { story: "main", blockId: paragraph.paraId, offset },
          },
        },
      });
      if (compiled.isErr()) throw compiled.error;
      const prepared = await prepareCanonicalDocxInput(
        await createDocx(compiled.value.document),
      );
      if (prepared.isErr()) throw prepared.error;
      for (const text of [REPLACEMENT, ""]) {
        for (const mode of ["direct", "tracked-changes"] as const) {
          const reviewer = await open(
            prepared.value.slice().buffer,
            "canonical",
          );
          const original = comments(reviewer);
          expect(original).toHaveLength(1);
          const before = reviewer.toDocument();
          const saved = await reviewer.toBuffer();
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [
                {
                  id: "replace",
                  type: "replaceBlock",
                  blockId: paragraph.paraId,
                  text,
                },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          if (result.applied.length === 0) {
            expect(result.skipped).toMatchObject([
              {
                reason: "unsupportedBlock",
                canonicalRefusal: { gap: "comment-replacement-alignment" },
              },
            ]);
            expect(reviewer.toDocument()).toEqual(before);
            expect(new Uint8Array(await reviewer.toBuffer())).toEqual(
              new Uint8Array(saved),
            );
            continue;
          }
          expect(result.skipped).toEqual([]);
          const pending = await reviewer.toBuffer();
          expect(reviewer.acceptAll().skipped).toEqual([]);
          expect(comments(reviewer)).toEqual(original);
          expect(
            comments(await open(await reviewer.toBuffer(), "canonical")),
          ).toEqual(original);
          if (mode === "tracked-changes") {
            const rejected = await open(pending, "canonical");
            expect(rejected.rejectAll().skipped).toEqual([]);
            expect(comments(rejected)).toEqual(original);
          }
        }
      }
    });
  }
  /**
   * A tracked replacement the new words cannot carry a comment into keeps
   * each comment on the words it deletes: rejecting gives every comment back,
   * and accepting keeps only comments that still cover words, reporting each
   * one it drops.
   */
  const expectTrackedReplacementKeepsComments = async ({
    reviewer,
    result,
    original,
  }: {
    reviewer: Awaited<ReturnType<typeof open>>;
    result: ReturnType<
      Awaited<ReturnType<typeof open>>["applyDocumentOperations"]
    >;
    original: ReturnType<typeof comments>;
  }) => {
    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["replace"]);
    const pending = await reviewer.toBuffer();
    const rejected = await open(pending, "canonical");
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(comments(rejected)).toEqual(original);
    expect(reviewer.acceptAll().skipped).toEqual([]);
    const kept = new Set(comments(reviewer).map(({ id }) => id));
    const retired = new Set(result.applied[0]?.retiredCommentIds ?? []);
    for (const { id } of original)
      expect(kept.has(id) || retired.has(id)).toBe(true);
  };
  for (const [name, scenario] of Object.entries(COMMENT_CASES)) {
    test(`${name}: supported replacements preserve threads; unsupported alignment refuses unchanged`, async () => {
      // Save updates package modification dates, independently of any edit.
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const prepared = await prepareCanonicalDocxInput(
        await createDocx(fromMarkdown(SOURCE)),
      );
      if (prepared.isErr()) throw prepared.error;
      // The legacy public producer creates source threads, including replies;
      // every editor opens the same source package and existing comment IDs.
      const producer = await open(prepared.value.slice().buffer, "legacy");
      const block = producer.getContent().at(0);
      if (!block) throw new Error("Source paragraph is missing");
      for (const [index, quote] of scenario.quotes.entries()) {
        apply(
          producer,
          {
            id: `comment-${index}`,
            type: "commentOnBlock",
            blockId: block.id,
            quote,
            comment: { text: `Note ${index}` },
          },
          "direct",
        );
        const comment = producer.getComments().at(-1);
        if (!comment) throw new Error("Produced comment is missing");
        expect(
          producer.replyTo(comment, { text: `Reply ${index}` }),
        ).not.toBeNull();
      }
      const bytes = await producer.toBuffer();
      const operation = {
        id: "replace",
        type: "replaceBlock",
        blockId: block.id,
        text: REPLACEMENT,
      } as const;
      const direct = await open(bytes, "canonical");
      const tracked = await open(bytes, "canonical");
      const original = comments(tracked);
      expect(original).toHaveLength(scenario.quotes.length);
      if (scenario.type === "alignment") {
        for (const mode of ["direct", "tracked-changes"] as const) {
          const reviewer = await open(bytes, "canonical");
          const before = reviewer.toDocument();
          const saved = await reviewer.toBuffer();
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [operation],
            },
            { undefinedReferences: "refuse" },
          );
          if (mode === "tracked-changes") {
            await expectTrackedReplacementKeepsComments({
              reviewer,
              result,
              original,
            });
            continue;
          }
          expect(result.applied).toEqual([]);
          expect(result.undoHandle).toBeNull();
          expect(result.skipped).toMatchObject([
            {
              id: "replace",
              reason: "unsupportedBlock",
              canonicalRefusal: { gap: "comment-replacement-alignment" },
            },
          ]);
          expect(reviewer.toDocument()).toEqual(before);
          expect(comments(reviewer)).toEqual(original);
          expect(new Uint8Array(await reviewer.toBuffer())).toEqual(
            new Uint8Array(saved),
          );
        }
        return;
      }
      apply(direct, operation, "direct");
      const result = apply(tracked, operation, "tracked-changes");
      const pending = await tracked.toBuffer();
      expect(tracked.acceptAll().skipped).toEqual([]);
      // The replacement takes every word these comments cover, and the
      // comments go with their words: the new words take no comment's range.
      const replaced: ReturnType<typeof comments> = [];
      expect(comments(direct)).toEqual(replaced);
      expect(comments(tracked)).toEqual(replaced);
      for (const reviewer of [direct, tracked])
        expect(
          comments(await open(await reviewer.toBuffer(), "canonical")),
        ).toEqual(replaced);
      const rejected = await open(pending, "canonical");
      expect(rejected.rejectAll().skipped).toEqual([]);
      expect(comments(rejected)).toEqual(original);
      for (const decision of ["accept", "reject"] as const) {
        for (const order of ["forward", "reverse"] as const) {
          const individually = await open(pending, "canonical");
          const changes = individually.getChanges();
          for (const change of order === "forward"
            ? changes
            : changes.toReversed())
            expect(
              decision === "accept"
                ? individually.acceptChange(change)
                : individually.rejectChange(change),
            ).toBe(true);
          const expected = decision === "accept" ? replaced : original;
          expect(comments(individually)).toEqual(expected);
          expect(
            comments(await open(await individually.toBuffer(), "canonical")),
          ).toEqual(expected);
        }
      }
      const undone = await open(bytes, "canonical");
      const undoResult = apply(undone, operation, "tracked-changes");
      if (undoResult.undoHandle === null || result.undoHandle === null)
        throw new Error("Replacement must have an undo handle");
      expect(undone.undoDocumentOperations(undoResult.undoHandle).status).toBe(
        "undone",
      );
      expect(comments(undone)).toEqual(original);
    });
  }
});
