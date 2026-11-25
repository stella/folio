import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_TYPES,
  type EditorIntent,
  type IdLease,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { readModel } from "../types/documentModel";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const fixture = async () => {
  const document = fromMarkdown("Left.\n\nRight.");
  const first = fromMarkdown(
    "| T1 |\n|---|\n| A1 |",
  ).package.document.content.find((block) => block.type === "table");
  const second = fromMarkdown(
    "| T2 |\n|---|\n| A2 |",
  ).package.document.content.find((block) => block.type === "table");
  if (first === undefined || second === undefined)
    panic("The producer fixture needs both tables.");
  document.package.document.content.splice(1, 0, first, second);
  return (await prepareCanonicalDocxInput(await createDocx(document)))
    .unwrap()
    .slice().buffer;
};

for (const leasing of ["unleased", "leased"] as const)
  for (const mode of ["editing", "suggesting"] as const)
    test(`${leasing} ${mode}: a public pending-table producer saves, resolves, and undoes table-side paragraph insertion`, async () => {
      const input = await fixture();
      const original = await FolioDocxReviewer.fromBuffer(input, {
        session: "canonical",
      });
      const initialTexts = original.getContent().map(({ text }) => text);
      const lease =
        leasing === "leased"
          ? ({
              blocks: [
                { kind: "paragraph", start: 0x40000001, end: 0x40000041 },
                { kind: "revision", start: 10000, end: 10400 },
                { kind: "control", start: 20000, end: 20400 },
              ],
            } as const satisfies IdLease)
          : undefined;
      const session = createCanonicalSession(
        await parseDocx(input),
        undefined,
        lease === undefined ? {} : { idLease: lease },
      ).unwrap();
      let state = EditorState.create({ doc: session.projection.doc });
      const publish = (commit: CanonicalCommit) => {
        const published = publishCanonicalProjection({
          session,
          state,
          commit,
        });
        if (published.isErr()) return false;
        state = published.value.state;
        return true;
      };
      const id = (text: string) =>
        createFolioAIEditSnapshot(state.doc).blocks.find(
          (block) => block.text === text,
        )?.id ?? panic("The parsed source needs its table anchor.");
      const firstTable = id("T1"),
        secondTable = id("T2");
      const operations = new CanonicalPublicOperations({
        undoLimit: Number.POSITIVE_INFINITY,
        session,
        getState: () => state,
        publish,
      });
      const pending = operations.apply({
        snapshot: createFolioAIEditSnapshot(state.doc),
        batch: {
          version: 1,
          mode: "tracked-changes",
          operations: [
            { id: "delete-first", type: "deleteTable", blockId: firstTable },
            { id: "delete-second", type: "deleteTable", blockId: secondTable },
          ],
        },
        undefinedReferences: "refuse",
        revisionStamp: { idSeed: 1, date: "2026-01-02T03:04:05Z" },
      });
      expect(pending.applied).toHaveLength(2);
      expect(pending.skipped).toEqual([]);
      const before = readModel(session.document),
        version = session.version;
      const allocate = createEditorIntentIdAllocator({
        reservedBlockIds: session.reservedBlockIds,
      });
      const table = fromMarkdown(
        "| Second |\n|---|\n| Value |",
      ).package.document.content.find((block) => block.type === "table");
      if (table === undefined)
        panic("The new table needs compiler-owned cells.");
      const tableIntent = {
        type: "tableBlock",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story: OP_STORIES.MAIN,
          blockId: secondTable,
          table,
          anchor: "table",
          side: "before",
        },
      } as const satisfies EditorIntent;
      const allocationLease = session.idLease;
      const leased =
        allocationLease === undefined ? {} : { idLease: allocationLease };
      const tableIds = allocate(before, tableIntent, allocationLease);
      const compileMode =
        mode === "editing"
          ? { type: "editing" as const, newIds: tableIds.newIds, ...leased }
          : {
              type: "suggesting" as const,
              newIds: tableIds.newIds,
              ...leased,
              revision: {
                id: tableIds.revisionId,
                author: "Reviewer",
                date: "2026-01-02T03:04:05Z",
              },
            };
      const compiledTable = compileEditorIntent(before, {
        intent: tableIntent,
        mode: compileMode,
      }).unwrap();
      const afterTable = applyDocumentOps(before, compiledTable.ops).unwrap()
        .document;
      const insertedTable = compiledTable.ops.find(
        (op) => op.type === DOCUMENT_OP_TYPES.INSERT_TABLE,
      );
      const tableAnchor = insertedTable?.table.rows
        .at(0)
        ?.cells.at(0)
        ?.content.at(0);
      if (tableAnchor?.type !== "paragraph" || tableAnchor.paraId === undefined)
        panic("The real table producer needs its owned anchor.");
      const paragraphIntent = {
        type: "insertParagraphs",
        story: OP_STORIES.MAIN,
        at: { type: "beforeTable", blockId: tableAnchor.paraId },
        paragraphs: [
          {
            type: "paragraph",
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "First words." }],
              },
            ],
          },
        ],
      } as const satisfies EditorIntent;
      const paragraphIds = allocate(
        afterTable,
        paragraphIntent,
        allocationLease,
      );
      const compiledParagraph = compileEditorIntent(afterTable, {
        intent: paragraphIntent,
        mode:
          mode === "editing"
            ? { type: "editing", newIds: paragraphIds.newIds, ...leased }
            : {
                type: "suggesting",
                ...leased,
                revision: {
                  id: paragraphIds.revisionId,
                  author: "Reviewer",
                  date: "2026-01-02T03:04:05Z",
                },
                newIds: paragraphIds.newIds,
              },
      }).unwrap();
      let events = 0;
      const unsubscribe = session.onCommitted(() => events++);
      const prepared = session
        .prepareOps(state, [...compiledTable.ops, ...compiledParagraph.ops], {
          type: "text",
          anchor: compiledParagraph.selection,
          head: compiledParagraph.selection,
        })
        .unwrap();
      expect(session.version).toBe(version);
      expect(session.document).toBe(before);
      expect(session.idLease).toBe(allocationLease);
      expect(events).toBe(0);
      expect(publish(prepared)).toBe(true);
      expect(events).toBe(1);
      const after = readModel(session.document);
      const saved = await createDocx(after);
      const accepted = await FolioDocxReviewer.fromBuffer(saved, {
        session: "canonical",
      });
      const acceptance = accepted.acceptAll();
      expect(acceptance.resolved).toBeGreaterThan(0);
      expect(acceptance.skipped).toEqual([]);
      expect(
        accepted
          .getContent()
          .map(({ text }) => text)
          .filter(Boolean),
      ).toEqual(["Left.", "First words.", "Second", "Value", "Right."]);
      if (mode === "suggesting") {
        const rejected = await FolioDocxReviewer.fromBuffer(saved, {
          session: "canonical",
        });
        const rejection = rejected.rejectAll();
        expect(rejection.resolved).toBeGreaterThan(0);
        expect(rejection.skipped).toEqual([]);
        expect(rejected.getContent().map(({ text }) => text)).toEqual(
          initialTexts,
        );
      }
      expect(
        publish(session.prepareUndo(state, OP_STORIES.MAIN).unwrap()),
      ).toBe(true);
      expect(readModel(session.document)).toStrictEqual(before);
      expect(
        publish(session.prepareRedo(state, OP_STORIES.MAIN).unwrap()),
      ).toBe(true);
      expect(readModel(session.document)).toStrictEqual(after);
      unsubscribe();
    });
