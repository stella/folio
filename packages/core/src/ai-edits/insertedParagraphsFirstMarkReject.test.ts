/**
 * Paragraphs inserted with tracking after the last paragraph of a container
 * add inserted paragraph marks. Rejecting only the first of them joins two
 * paragraphs, and every other change of the insertion stays pending: the
 * inserted words, the later inserted marks and the pending property changes.
 * Rejecting by revision (the change's card) and at the caret keeps them,
 * through a save and reopen, and what is left resolves.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { createHarnessState, findTextblock } from "../__tests__/editorHarness";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { paragraph, table } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import {
  findChangeAtPosition,
  rejectChange,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { FolioDocxReviewer } from "./headless";

type Container = "body" | "cell";

const seed = async (container: Container): Promise<ArrayBuffer> => {
  const document = fromMarkdown("Body end.");
  const anchor = paragraph("kept");
  anchor.formatting = { alignment: "center" };
  if (container === "body") document.package.document.content = [anchor];
  else {
    const cellTable = table({ rows: [["placeholder"]] });
    const cell =
      cellTable.rows.at(0)?.cells.at(0) ?? panic("Missing fixture cell");
    cell.content = [anchor];
    document.package.document.content = [cellTable, paragraph("Body end.")];
  }
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  return new Uint8Array(docx).buffer;
};

const open = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });

const insertAfterAnchor = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
) => {
  const anchor =
    reviewer.getContent().find(({ text }) => text === "kept") ??
    panic("Missing anchor block");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [
        {
          id: "paragraphs",
          type: "insertAfterBlock",
          blockId: anchor.id,
          text: "first\nsecond",
          alignment: "right",
          formattingScope: "allParagraphs",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.applied.map(({ id }) => id)).toEqual(["paragraphs"]);
};

const changes = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getChanges()
    .map(({ id, type }) => [id, type] as const)
    .toSorted(([left], [right]) => left - right);

const content = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ kind, text }) => ({ kind, text }));

// In the body the anchor's own mark is the inserted one and the last new
// paragraph holds the mark the anchor had, with a pending change of the
// anchor's properties. In a cell the new paragraphs' marks are the inserted
// ones. Either way the first inserted mark joins two paragraphs.
const EXPECTED = {
  body: {
    joinedFrom: "kept",
    joined: "keptfirst",
    rest: [
      "insertion",
      "insertion",
      "paragraphMarkInserted",
      "paragraphPropertiesChanged",
      "paragraphPropertiesChanged",
    ],
  },
  cell: {
    joinedFrom: "first",
    joined: "firstsecond",
    rest: ["insertion", "insertion", "paragraphMarkInserted"],
  },
} as const;

describe.each(["body", "cell"] as const)(
  "rejecting only the first inserted paragraph mark after a final paragraph (%s)",
  (container) => {
    test("keeps every other pending change, by revision, through save", async () => {
      const expected = EXPECTED[container];
      const original = await seed(container);
      const pending = await open(original);
      insertAfterAnchor(pending, "tracked-changes");
      const pendingBytes = await pending.toBuffer();
      const firstMark =
        pending
          .getChanges()
          .find(({ type }) => type === "paragraphMarkInserted") ??
        panic("Missing an inserted paragraph mark");
      expect(firstMark.text).toBe(expected.joinedFrom);
      const rest = changes(pending).filter(([id]) => id !== firstMark.id);
      expect(rest.map(([, type]) => type).toSorted()).toEqual([
        ...expected.rest,
      ]);

      // The change's card rejects by revision.
      const byRevision = await open(pendingBytes);
      expect(byRevision.rejectChange(firstMark.id)).toBe(true);
      expect(changes(byRevision)).toEqual(rest);
      const joined = content(byRevision);
      expect(joined.map(({ text }) => text)).toContain(expected.joined);

      // Saved and reopened, the rest is still pending and resolves both ways.
      const reopened = await open(await byRevision.toBuffer());
      expect(changes(reopened)).toEqual(rest);
      expect(content(reopened)).toEqual(joined);

      const rejecting = await open(await byRevision.toBuffer());
      rejecting.rejectAll();
      expect(changes(rejecting)).toEqual([]);
      expect(content(rejecting)).toEqual(content(await open(original)));

      const accepting = await open(await byRevision.toBuffer());
      accepting.acceptAll();
      expect(changes(accepting)).toEqual([]);
      expect(content(accepting)).toEqual(joined);
    });
  },
);

test("rejecting the anchor's mark at its caret in the body keeps every other pending change", async () => {
  const pending = await open(await seed("body"));
  insertAfterAnchor(pending, "tracked-changes");
  const pendingBytes = await pending.toBuffer();
  const anchorMark =
    pending.getChanges().find(({ type }) => type === "paragraphMarkInserted") ??
    panic("Missing the anchor's inserted mark");
  const rest = changes(pending).filter(([id]) => id !== anchorMark.id);
  const byRevision = await open(pendingBytes);
  expect(byRevision.rejectChange(anchorMark.id)).toBe(true);

  // The anchor's own words are not tracked, so the caret at its end reaches
  // only its paragraph mark.
  const parsed = await parseDocx(pendingBytes);
  let state = createHarnessState(parsed, "editing");
  const markBlock =
    findTextblock(state.doc, "kept") ?? panic("Missing the anchor textblock");
  const caret = markBlock.pos + markBlock.node.nodeSize - 1;
  const range = findChangeAtPosition(state, caret, caret);
  expect(
    rejectChange(range.from, range.to)(state, (transaction) => {
      state = state.apply(transaction);
    }),
  ).toBe(true);
  const atCaret = await open(
    await createDocx(
      fromProseDoc(state.doc, parsed, {
        stylesheetSource: { type: "package" },
      }),
    ),
  );
  expect(changes(atCaret)).toEqual(rest);
  expect(content(atCaret)).toEqual(content(byRevision));
});
