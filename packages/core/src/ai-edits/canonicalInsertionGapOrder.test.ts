import { panic } from "better-result";
import { expect, test } from "bun:test";

import { storyParagraphs } from "@stll/docx-core/ops";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";
import { supportsTableInsertionMode } from "./tableInsertionModes";

const KINDS = ["paragraph", "paragraphs", "table", "signature"] as const;
type Kind = (typeof KINDS)[number];
type InsertionOptions = {
  kind: Kind;
  label: string;
  blockId: string;
  side: "before" | "after";
};
const insertion = ({
  kind,
  label,
  blockId,
  side,
}: InsertionOptions): FolioDocumentOperation => {
  switch (kind) {
    case "paragraph":
    case "paragraphs":
      return {
        id: label,
        type: side === "before" ? "insertBeforeBlock" : "insertAfterBlock",
        blockId,
        text:
          kind === "paragraph"
            ? `${label} words.`
            : `${label} first.\n${label} second.`,
      };
    case "table":
      return {
        id: label,
        type: "insertTable",
        blockId,
        position: side,
        rows: [[`${label} term`, `${label} value`]],
      };
    case "signature":
      return {
        id: label,
        type: "insertSignatureTable",
        blockId,
        position: side,
        parties: [{ name: `${label} signer` }],
      };
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

const labels = (reviewer: FolioDocxReviewer) => {
  const result: string[] = [];
  for (const { text } of reviewer.getContent()) {
    const label = /\b(?:First|Second|Third|Left|Right)\b/u.exec(text)?.[0];
    if (label !== undefined && label !== result.at(-1)) result.push(label);
  }
  return result;
};

for (const mode of ["direct", "tracked-changes"] as const) {
  const kinds = KINDS.filter(
    (kind) =>
      kind !== "signature" ||
      supportsTableInsertionMode("insertSignatureTable", mode),
  );
  for (const firstKind of kinds)
    for (const secondKind of kinds)
      for (const sides of ["before-after", "after-before"] as const)
        test(`${mode} ${firstKind}/${secondKind} ${sides}: a shared gap retains input order`, async () => {
          const bytes = (
            await prepareCanonicalDocxInput(
              await createDocx(fromMarkdown("Left.\n\nRight.")),
            )
          )
            .unwrap()
            .slice().buffer;
          const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
            session: "canonical",
          });
          const left = reviewer
            .getContent()
            .find(({ text }) => text === "Left.");
          const right = reviewer
            .getContent()
            .find(({ text }) => text === "Right.");
          if (!left || !right)
            panic("The parsed fixture needs both gap anchors.");
          const firstBefore = sides === "before-after";
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [
                insertion({
                  kind: firstKind,
                  label: "First",
                  blockId: firstBefore ? right.id : left.id,
                  side: firstBefore ? "before" : "after",
                }),
                insertion({
                  kind: secondKind,
                  label: "Second",
                  blockId: firstBefore ? left.id : right.id,
                  side: firstBefore ? "after" : "before",
                }),
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied).toHaveLength(2);
          expect(result.skipped).toEqual([]);
          if (mode === "tracked-changes") {
            const accepted = reviewer.acceptAll();
            expect(accepted.resolved).toBeGreaterThan(0);
            expect(accepted.skipped).toEqual([]);
            expect(reviewer.getChanges()).toEqual([]);
          }
          expect(labels(reviewer)).toEqual([
            "Left",
            "First",
            "Second",
            "Right",
          ]);
          const reopened = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
            { session: "canonical" },
          );
          expect(labels(reopened)).toEqual([
            "Left",
            "First",
            "Second",
            "Right",
          ]);
        });
}

const TERMINAL_ORDERS = [
  ["paragraph", "table", "paragraphs"],
  ["paragraph", "paragraphs", "table"],
  ["table", "paragraph", "paragraphs"],
  ["table", "paragraphs", "paragraph"],
  ["paragraphs", "paragraph", "table"],
  ["paragraphs", "table", "paragraph"],
] as const satisfies readonly (readonly Kind[])[];
for (const mode of ["direct", "tracked-changes"] as const)
  for (const kinds of TERMINAL_ORDERS)
    test(`${mode} terminal ${kinds.join("/")}: normal admission preserves a legal tail, undo and resolution`, async () => {
      const bytes = (
        await prepareCanonicalDocxInput(await createDocx(fromMarkdown("Left.")))
      )
        .unwrap()
        .slice().buffer;
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
        session: "canonical",
      });
      const initial = reviewer.getContent();
      const left = initial.find(({ text }) => text === "Left.");
      if (left === undefined)
        panic("The terminal fixture needs its source paragraph.");
      const operations = kinds.map((kind, index) =>
        insertion({
          kind,
          label:
            ["First", "Second", "Third"][index] ??
            panic("The terminal matrix needs a label."),
          blockId: left.id,
          side: "after",
        }),
      );
      // Normal admission runs these same-anchor operations from the last
      // input backwards. A tracked table takes the source mark, so earlier
      // paragraph insertions still meet the native pending-mark refusal.
      // Reordering admitted operations must not expand that capability.
      const firstAdmitted = mode === "direct" ? 0 : kinds.indexOf("table");
      const admittedIds = operations.slice(firstAdmitted).map(({ id }) => id);
      const declinedIds = operations
        .slice(0, firstAdmitted)
        .map(({ id }) => id);
      const expectedLabels = ["Left", ...admittedIds];
      const apply = () => {
        const edited = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations,
          },
          { undefinedReferences: "refuse" },
        );
        expect(edited.applied.map(({ id }) => id)).toEqual(admittedIds);
        expect(edited.skipped).toEqual(
          declinedIds.map((id) => ({
            id,
            reason: "unsupportedBlock",
            message:
              "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of review records.",
            canonicalRefusal: { gap: "publicOps.unsupportedInline" },
          })),
        );
        expect(edited.undoHandle).not.toBeNull();
        return (
          edited.undoHandle ??
          panic("The applied batch needs its inverse handle.")
        );
      };
      const undo = apply();
      expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
      expect(reviewer.getContent()).toEqual(initial);
      apply();
      if (mode === "tracked-changes") {
        const rejected = reviewer.rejectAll();
        expect(rejected.resolved).toBeGreaterThan(0);
        expect(rejected.skipped).toEqual([]);
        expect(reviewer.getContent()).toEqual(initial);
        apply();
        const accepted = reviewer.acceptAll();
        expect(accepted.resolved).toBeGreaterThan(0);
        expect(accepted.skipped).toEqual([]);
        expect(reviewer.getChanges()).toEqual([]);
      }
      expect(labels(reviewer)).toEqual(expectedLabels);
      const content = reviewer.getContent();
      expect(new Set(content.map(({ id }) => id)).size).toBe(content.length);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session: "canonical" },
      );
      expect(labels(reopened)).toEqual(expectedLabels);
    });

test("an interleaved source-table comment retains its admission, paragraph IDs and revision clock", async () => {
  const bytes = (
    await prepareCanonicalDocxInput(
      await createDocx(
        fromMarkdown("Left.\n\n| T | U |\n|---|---|\n| V | W |\n\nRight."),
      ),
    )
  )
    .unwrap()
    .slice().buffer;
  const open = () =>
    FolioDocxReviewer.fromBuffer(bytes, { session: "canonical" });
  const normal = await open(),
    planned = await open();
  const find = (text: string) =>
    normal.getContent().find((block) => block.text === text) ??
    panic("The source fixture lost a paragraph.");
  const a = find("Left.").id,
    b = find("Right.").id,
    t = find("T").id;
  const date = "2026-01-02T03:04:05Z";
  const apply = (
    reviewer: FolioDocxReviewer,
    operations: readonly FolioDocumentOperation[],
    idSeed: number,
  ) => {
    const edited = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations,
      },
      { undefinedReferences: "refuse", revisionStamp: { idSeed, date } },
    );
    expect(edited.applied).toHaveLength(operations.length);
    expect(edited.skipped).toEqual([]);
    return edited;
  };
  for (const reviewer of [normal, planned])
    apply(reviewer, [{ id: "delete", type: "deleteTable", blockId: t }], 1);
  const before = {
    id: "before",
    type: "insertBeforeBlock",
    blockId: b,
    text: "Second words.",
  } as const satisfies FolioDocumentOperation;
  const comment = {
    id: "comment",
    type: "commentOnBlock",
    blockId: t,
    comment: { text: "Source table" },
  } as const satisfies FolioDocumentOperation;
  const after = {
    id: "after",
    type: "insertAfterBlock",
    blockId: a,
    text: "First words.",
  } as const satisfies FolioDocumentOperation;
  // The shared scheduler admits source annotations before content changes.
  const commentResult = apply(normal, [comment], 100);
  const beforeResult = apply(normal, [before], 100);
  const afterResult = apply(
    normal,
    [after],
    Math.max(
      ...beforeResult.applied.flatMap(({ revisionIds }) => revisionIds ?? []),
    ) + 1,
  );
  const actual = apply(planned, [after, comment, before], 100);
  expect(actual.applied).toEqual([
    ...afterResult.applied,
    ...commentResult.applied,
    ...beforeResult.applied,
  ]);
  const owned = (reviewer: FolioDocxReviewer) =>
    reviewer
      .getContent()
      .filter(({ text }) => text === "First words." || text === "Second words.")
      .map(({ text, id }) => ({ text, id }))
      .toSorted((left, right) => left.text.localeCompare(right.text));
  expect(owned(planned)).toEqual(owned(normal));
  expect(planned.getComments().map(({ id }) => id)).toEqual(
    normal.getComments().map(({ id }) => id),
  );
});

test("a table cell keeps its admitted identity when the group no longer needs its terminal carrier", async () => {
  const bytes = (
    await prepareCanonicalDocxInput(await createDocx(fromMarkdown("Left.")))
  )
    .unwrap()
    .slice().buffer;
  const normal = await FolioDocxReviewer.fromBuffer(bytes, {
    session: "canonical",
  });
  const planned = await FolioDocxReviewer.fromBuffer(bytes, {
    session: "canonical",
  });
  const left =
    normal.getContent().find(({ text }) => text === "Left.") ??
    panic("The source anchor is missing.");
  const table = insertion({
    kind: "table",
    label: "First",
    blockId: left.id,
    side: "after",
  });
  const paragraph = insertion({
    kind: "paragraph",
    label: "Second",
    blockId: left.id,
    side: "after",
  });
  const apply = (
    reviewer: FolioDocxReviewer,
    operations: readonly FolioDocumentOperation[],
  ) => {
    const edited = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations,
      },
      { undefinedReferences: "refuse" },
    );
    expect(edited.applied).toHaveLength(operations.length);
    expect(edited.skipped).toEqual([]);
  };
  // Equal-position content is admitted in reverse input order. The normal
  // table insertion needs a terminal carrier that replay no longer needs.
  apply(normal, [table]);
  apply(normal, [paragraph]);
  apply(planned, [paragraph, table]);
  const owned = (reviewer: FolioDocxReviewer) =>
    reviewer
      .getContent()
      .filter(({ text }) => text !== "")
      .map(({ id, text }) => ({ id, text }))
      .toSorted((a, b) => a.text.localeCompare(b.text));
  expect(owned(planned)).toEqual(owned(normal));
});

for (const mode of ["direct", "tracked-changes"] as const)
  for (const order of ["paragraph-first", "table-first"] as const)
    test(`${mode} ${order}: a shared gap between pending tables preserves its physical source boundary`, async () => {
      const source = fromMarkdown("Left.\n\nRight.");
      const tables = ["One", "Two"].map((name) => {
        const table = fromMarkdown(
          `| ${name} | Value |\n|---|---|\n| Term | Detail |`,
        ).package.document.content.find((block) => block.type === "table");
        return table ?? panic("The source fixture needs a table.");
      });
      source.package.document.content.splice(1, 0, ...tables);
      const bytes = (await prepareCanonicalDocxInput(await createDocx(source)))
        .unwrap()
        .slice().buffer;
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
        session: "canonical",
      });
      const original = reviewer.getContent();
      const id = (text: string) =>
        original.find((block) => block.text === text)?.id ??
        panic("The source fixture lost its anchor.");
      const pending = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: ["One", "Two"].map((text) => ({
            id: text,
            type: "deleteTable",
            blockId: id(text),
          })),
        },
        { undefinedReferences: "refuse" },
      );
      expect(pending.applied).toHaveLength(2);
      expect(pending.skipped).toEqual([]);
      const before = reviewer.toDocument();
      const paragraphLabel = order === "paragraph-first" ? "First" : "Second";
      const paragraph = {
        id: paragraphLabel,
        type: "insertAfterBlock",
        blockId: id("Left."),
        text: `${paragraphLabel} words.`,
        comment: { text: "Inserted clause" },
      } as const satisfies FolioDocumentOperation;
      const table = insertion({
        kind: "table",
        label: order === "table-first" ? "First" : "Second",
        blockId: id("Two"),
        side: "before",
      });
      const operations =
        order === "paragraph-first" ? [paragraph, table] : [table, paragraph];
      const apply = () => {
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations,
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.applied).toHaveLength(2);
        expect(result.skipped).toEqual([]);
        expect(labels(reviewer)).toEqual(["Left", "First", "Second", "Right"]);
        const inserted = reviewer
          .getContent()
          .find(({ text }) => text === paragraph.text);
        expect(inserted).toBeDefined();
        expect(
          reviewer.getComments().map(({ text, anchoredText, blockId }) => ({
            text,
            anchoredText,
            blockId,
          })),
        ).toEqual([
          {
            text: "Inserted clause",
            anchoredText: paragraph.text,
            blockId: inserted?.id,
          },
        ]);
        const content = reviewer.toDocument().package.document.content;
        const tableIndex = (name: string) =>
          content.findIndex(
            (block) =>
              block.type === "table" &&
              storyParagraphs({ content: [block] }).some(
                ({ paragraph: item }) => item.paraId === id(name),
              ),
          );
        const one = tableIndex("One"),
          two = tableIndex("Two");
        expect(one).toBeGreaterThanOrEqual(0);
        expect(two).toBe(one + 3);
        expect(content.slice(one + 1, two).map(({ type }) => type)).toEqual(
          order === "paragraph-first"
            ? ["paragraph", "table"]
            : ["table", "paragraph"],
        );
        return result.undoHandle ?? panic("The applied gap needs its inverse.");
      };
      expect(reviewer.undoDocumentOperations(apply()).status).toBe("undone");
      expect(reviewer.toDocument()).toStrictEqual(before);
      expect(reviewer.getComments()).toEqual([]);
      apply();
      const comments = (current: FolioDocxReviewer) =>
        current
          .getComments()
          .map(({ id: commentId, text, anchoredText, blockId }) => ({
            id: commentId,
            text,
            anchoredText,
            blockId,
          }));
      const insertedComments = comments(reviewer);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        {
          session: "canonical",
        },
      );
      expect(labels(reopened)).toEqual(["Left", "First", "Second", "Right"]);
      expect(comments(reopened)).toEqual(insertedComments);
      const accepted = reopened.acceptAll();
      expect(accepted.skipped).toEqual([]);
      expect(reopened.getChanges()).toEqual([]);
      expect(comments(reopened)).toEqual(insertedComments);
      expect(labels(reopened)).toEqual(["Left", "First", "Second", "Right"]);
      if (mode === "tracked-changes") {
        const rejected = reviewer.rejectAll();
        expect(rejected.skipped).toEqual([]);
        expect(reviewer.getChanges()).toEqual([]);
        expect(reviewer.getContent()).toEqual(original);
        expect(reviewer.getComments()).toEqual([]);
      }
    });
