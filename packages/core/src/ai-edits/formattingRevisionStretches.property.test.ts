import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown";
import { expectRunPropertyChangeMarkAttrs } from "../prosemirror/attrs";
import { schema } from "../prosemirror/schema";
import { FolioDocxReviewer } from "./headless";
import { separateRevisionStretches } from "./revisionStretches";
import { createFolioAITextRangeHandle } from "./snapshot";

test(
  "formatting carrier allocation is unique and reaches a fixed point",
  () => {
    assertProperty(
      fc.property(fc.integer({ min: 2, max: 12 }), (count) => {
        const formatting = schema.mark("runPropertyChange", {
          changes: [
            { type: "runPropertyChange", info: { id: 1, author: "Reviewer" } },
          ],
        });
        const nodes = Array.from({ length: count }, (_, index) =>
          index % 2 === 0
            ? schema.text("Clause", [formatting])
            : schema.node("tab", null, null, [formatting]),
        );
        const tr = EditorState.create({
          doc: schema.node("doc", null, [
            schema.node("paragraph", null, nodes),
          ]),
        }).tr;
        const separated = separateRevisionStretches({ tr, revisionSeed: 2 });
        const ids: number[] = [];
        tr.doc.descendants((node) => {
          const mark = node.marks.find(
            ({ type }) => type.name === "runPropertyChange",
          );
          if (mark)
            ids.push(
              ...expectRunPropertyChangeMarkAttrs(mark).changes.map(
                ({ info }) => info.id,
              ),
            );
        });
        expect(ids).toHaveLength(count);
        expect(new Set(ids).size).toBe(count);
        const normalized = tr.doc;
        const repeated = separateRevisionStretches({
          tr,
          revisionSeed: separated.nextRevisionId,
        });
        expect(repeated.minted).toEqual([]);
        expect(tr.doc.eq(normalized)).toBe(true);
      }),
      {
        id: "formatting carrier allocation is unique and reaches a fixed point",
        numRuns: 25,
      },
    );
  },
  propertyTestTimeout(10_000),
);

const CASES = [
  "Alpha Bravo Omega",
  "Alpha 👍🏽 Omega",
  "Článek 日本語 Omega",
].flatMap((text) =>
  (["none", "inside"] as const).map((comment) => ({ text, comment })),
);

const formattingDocument = (text: string) => {
  const document = fromMarkdown(text);
  for (const [index, block] of document.package.document.content.entries()) {
    if (block.type !== "paragraph")
      panic("The formatting fixture must contain paragraphs");
    // Authored paragraph identities stay stable when their text changes on save.
    block.paraId = (0x34000001 + index).toString(16).toUpperCase();
  }
  return document;
};

// Existing formatting generators tested formatting once, without a later edit
// splitting its run or interleaving a comment with the formatting revision.
test.each(CASES)(
  "formatting split by a replacement and $comment comment saves the same changes ($text)",
  async ({ text, comment }) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(formattingDocument(text)),
      {
        author: "Reviewer",
      },
    );
    const block = reviewer.snapshot().blocks.at(0);
    if (!block || block.kind === "diagnostic")
      panic("The formatting fixture lost its paragraph");
    const formattingRange = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 0,
      endOffset: text.length,
    });
    if (!formattingRange) panic("The formatting fixture lost its text range");
    const apply = (
      operations: Parameters<
        typeof reviewer.applyDocumentOperations
      >[0]["operations"],
    ) => {
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations,
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(operations.length);
      expect(result.skipped).toEqual([]);
    };
    apply([
      {
        id: "format",
        type: "formatRange",
        range: formattingRange,
        formatting: { italic: true },
      },
    ]);
    const start = text.lastIndexOf("Omega");
    if (comment === "inside") {
      const commentRange = createFolioAITextRangeHandle({
        blockId: block.id,
        text: block.text,
        startOffset: start,
        endOffset: start + 5,
      });
      if (!commentRange) panic("The comment fixture lost its text range");
      apply([
        {
          id: "comment",
          type: "commentOnRange",
          range: commentRange,
          comment: { text: "Review this term." },
        },
      ]);
    }
    apply([
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: block.id,
        find: "Omega",
        replace: "Revised",
      },
    ]);
    const before = reviewer.getChanges();
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
    );
    expect(reopened.getChanges()).toEqual(before);
    expect(
      (
        await FolioDocxReviewer.fromBuffer(await reopened.toBuffer())
      ).getChanges(),
    ).toEqual(before);
  },
  propertyTestTimeout(60_000),
);

test(
  "formatting interleaved with a commented deletion preserves its logical text group",
  async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(
        formattingDocument("Alpha Bravo Omega\n\nClosing clause."),
      ),
      { author: "Reviewer" },
    );
    const block = reviewer.snapshot().blocks.at(0);
    if (!block || block.kind === "diagnostic")
      panic("The formatting fixture lost its paragraph");
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 6,
      endOffset: 11,
    });
    if (!range) panic("The formatting fixture lost its text range");
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "comment",
            type: "commentOnRange",
            range,
            comment: { text: "Review this term." },
          },
          {
            id: "format",
            type: "formatRange",
            range,
            formatting: { italic: true },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(2);
    expect(
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            { id: "delete", type: "deleteBlock", blockId: block.id },
          ],
        },
        { undefinedReferences: "refuse" },
      ).applied,
    ).toHaveLength(1);
    const before = reviewer.getChanges();
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
    );
    expect(reopened.getChanges()).toEqual(before);
    const deletion = reopened
      .getChanges()
      .find(({ type }) => type === "deletion");
    if (!deletion) panic("The commented deletion lost its logical change");
    expect(deletion.text).toBe("Alpha Bravo Omega");
    expect(reopened.rejectChange(deletion.id)).toBe(true);
    expect(reopened.getContent().at(0)?.text).toBe("Alpha Bravo Omega");
  },
  propertyTestTimeout(60_000),
);
