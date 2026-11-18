import { panic } from "better-result";
import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import type { Document, Paragraph } from "@stll/docx-core/model";
import {
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  type IdLease,
  OP_STORIES,
  REVISION_DECISIONS,
  storyParagraphs,
} from "@stll/docx-core/ops";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { readModel } from "../types/documentModel";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

// Compare snapshot bytes with the writer's modification timestamp held fixed.
beforeAll(() => setSystemTime(new Date("2026-01-02T03:04:05Z")));
afterAll(() => setSystemTime());

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const sourceFacts = (document: Document, ids: readonly string[]) => ({
  paragraphs: storyParagraphs(document.package.document)
    .filter(({ paragraph: item }) => ids.includes(item.paraId ?? ""))
    .map(({ paragraph: item }) => ({
      paraId: item.paraId,
      content: item.content,
      formatting: item.formatting,
    })),
  comments: document.package.document.comments?.map(
    ({ id, parentId, author, content }) => ({ id, parentId, author, content }),
  ),
});

for (const side of ["beforeTable", "afterTable"] as const)
  for (const order of ["mark-first", "table-first"] as const)
    test(`${side} ${order}: refusing a new table-side owner preserves source identities and threads through selective review and saves`, async () => {
      const bytes = (
        await prepareCanonicalDocxInput(
          await createDocx(fromMarkdown("Source")),
        )
      )
        .unwrap()
        .slice().buffer;
      let session = createCanonicalSession(await parseDocx(bytes)).unwrap();
      let state = EditorState.create({ doc: session.projection.doc });
      const first = storyParagraphs(
        readModel(session.document).package.document,
      ).at(0)?.paragraph;
      const sourceId =
        first?.paraId ??
        panic("The parsed producer needs its source paragraph.");
      const position = {
        story: OP_STORIES.MAIN,
        blockId: sourceId,
        offset: 0,
      } as const;
      const commit = (ops: readonly DocumentOp[]) => {
        const prepared = session
          .prepareOps(state, ops, {
            type: "text",
            anchor: position,
            head: position,
          })
          .unwrap();
        state = publishCanonicalProjection({
          session,
          state,
          commit: prepared,
        }).unwrap().state;
      };
      const compile = (intent: EditorIntent) => {
        const before = readModel(session.document);
        const idLease =
          session.idLease ??
          panic(
            "The source-ownership proof needs its finite allocation lease.",
          );
        const ids = createEditorIntentIdAllocator({
          reservedBlockIds: session.reservedBlockIds,
        })(before, intent, idLease);
        return compileEditorIntent(before, {
          intent,
          mode: {
            type: "suggesting",
            revision: { id: ids.revisionId, author: "Reviewer" },
            idLease,
            newIds: ids.newIds,
          },
        }).unwrap();
      };
      commit([
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: OP_STORIES.MAIN,
          blockId: sourceId,
          patch: { alignment: "center" },
        },
        {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: {
            id: 7,
            author: "Reviewer",
            content: [paragraph("20000001", "Source comment")],
          },
          anchor: {
            kind: "range",
            from: position,
            to: { ...position, offset: 3 },
          },
        },
        {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: {
            id: 8,
            author: "Reviewer",
            content: [paragraph("20000002", "Reply")],
          },
          anchor: { kind: "reply", parentId: 7 },
        },
      ]);
      // Reopen the actual authored comment package before the later edits;
      // save baselines then own the serialized comment-part identifiers.
      const commentSource = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
      });
      const idLease = {
        blocks: [
          { kind: "paragraph", start: 0x40000001, end: 0x40000041 },
          { kind: "revision", start: 10000, end: 10400 },
          { kind: "control", start: 20000, end: 20400 },
        ],
      } as const satisfies IdLease;
      session = createCanonicalSession(
        await parseDocx(commentSource.buffer),
        undefined,
        { idLease },
      ).unwrap();
      state = EditorState.create({ doc: session.projection.doc });
      const sibling = compile({
        type: "insertParagraphs",
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: sourceId },
        paragraphs: [
          {
            type: "paragraph",
            content: paragraph("00000000", "Standing sibling").content,
          },
        ],
      });
      commit(sibling.ops);
      const sourceMark = storyParagraphs(
        readModel(session.document).package.document,
      ).find(({ paragraph: item }) => item.paraId === sourceId)?.paragraph
        .pPrMark;
      if (sourceMark?.survivor !== "first")
        panic("The actual source producer needs its first-owner mark.");
      const siblingParagraph = storyParagraphs(
        readModel(session.document).package.document,
      ).find(
        ({ paragraph: item }) => item.paraId === sibling.selection.blockId,
      )?.paragraph;
      const contentRevision = siblingParagraph?.content.find(
        (node) => node.type === "insertion",
      );
      if (contentRevision?.type !== "insertion")
        panic("The actual sibling producer needs its content revision.");
      commit([
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [contentRevision.info.id],
          decision: REVISION_DECISIONS.ACCEPT,
        },
      ]);
      const ids = [sourceId, sibling.selection.blockId];
      const beforeTable = readModel(session.document);
      const expected = sourceFacts(beforeTable, ids);
      const table = compile({
        type: "tableBlock",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story: OP_STORIES.MAIN,
          blockId: sourceId,
          side: "after",
          table: {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [paragraph("30000011", "Cell")],
                  },
                ],
              },
            ],
          },
        },
      });
      const tableRevisionIds = applyDocumentOps(
        readModel(session.document),
        table.ops,
      ).unwrap().revisions;
      commit(table.ops);
      const before = readModel(session.document);
      const version = session.version;
      const lease = session.idLease;
      const reserved = [...session.reservedBlockIds];
      expect(session.beginSaveBaseline(version)).toBe(true);
      const savedBefore = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
      });
      expect(
        session.completeSaveBaseline(
          version,
          await parseDocx(savedBefore.buffer),
        ),
      ).toBe(true);
      const attempted = compile({
        type: "insertParagraphs",
        story: OP_STORIES.MAIN,
        at: { type: side, blockId: table.selection.blockId },
        paragraphs: [{ type: "paragraph", content: [] }],
      });
      let events = 0;
      const unsubscribe = session.onCommitted(() => events++);
      const refused = session.prepareOps(state, attempted.ops, {
        type: "text",
        anchor: attempted.selection,
        head: attempted.selection,
      });
      expect(refused.isErr()).toBe(true);
      expect(readModel(session.document)).toBe(before);
      expect(session.version).toBe(version);
      expect(session.idLease).toBe(lease);
      expect([...session.reservedBlockIds]).toEqual(reserved);
      expect(events).toBe(0);
      const savedAfter = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
      });
      expect(savedAfter.buffer).toEqual(savedBefore.buffer);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state, OP_STORIES.MAIN).unwrap(),
      }).unwrap().state;
      expect(readModel(session.document)).toStrictEqual(beforeTable);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareRedo(state, OP_STORIES.MAIN).unwrap(),
      }).unwrap().state;
      expect(readModel(session.document)).toStrictEqual(before);
      expect(events).toBe(2);
      events = 0;
      const decisions = {
        mark: {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [sourceMark.info.id],
          decision: REVISION_DECISIONS.ACCEPT,
        },
        table: {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [...tableRevisionIds],
          decision: REVISION_DECISIONS.REJECT,
        },
      } as const satisfies Record<string, DocumentOp>;
      for (const op of order === "mark-first"
        ? [decisions.mark, decisions.table]
        : [decisions.table, decisions.mark]) {
        commit([op]);
        const saved = await serializeCanonicalSave({
          snapshot: session.captureSaveSnapshot(),
        });
        expect(sourceFacts(await parseDocx(saved.buffer), ids)).toStrictEqual(
          sourceFacts(readModel(session.document), ids),
        );
      }
      expect(sourceFacts(readModel(session.document), ids)).toStrictEqual(
        expected,
      );
      expect(events).toBe(2);
      unsubscribe();
    });
