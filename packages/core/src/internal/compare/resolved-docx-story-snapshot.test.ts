import { describe, expect, test } from "bun:test";

import {
  numberingReferenceKeysOf,
  sourceDocumentOf,
  storyTablesOf,
} from "../../ai-edits/snapshot";
import type { FolioContentPropertySet } from "../../compare/content-types";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Paragraph, Table, TableCell } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxAuthoredRunsForBlock,
  resolvedDocxContentBlocks,
  resolvedDocxHasExactAuthoredRuns,
  resolvedDocxOperationSnapshot,
  resolvedDocxSourceDocument,
  resolvedDocxTableNodes,
} from "./resolved-docx-story-snapshot";

const property = (properties: FolioContentPropertySet, key: string) =>
  properties.find((candidate) => candidate.key === key)?.value;

const styledDocument = () => {
  const document = createEmptyDocument();
  document.package.styles = {
    docDefaults: { rPr: { fontSize: 22 } },
    styles: [
      {
        type: "paragraph",
        styleId: "Clause",
        pPr: { alignment: "center" },
        rPr: { bold: true },
      },
    ],
  };
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "A1000001",
    formatting: { styleId: "Clause", alignment: "right" },
    content: [
      {
        type: "run",
        formatting: { italic: true },
        content: [{ type: "text", text: "Alpha" }],
      },
    ],
  };
  document.package.document.content = [paragraph];
  return { document, paragraph };
};

describe("owned live DOCX story projection", () => {
  test("uses the live style cascade once and owns nested paragraph and run state", () => {
    const { document, paragraph } = styledDocument();
    const source = toProseDoc(document);
    const snapshot = createResolvedDocxStorySnapshot({
      document,
      story: { type: "main" },
      sourceDocument: source,
    });
    if (!snapshot) throw new Error("main story projection missing");
    const block = resolvedDocxContentBlocks(snapshot).at(0);
    if (!block) throw new Error("projected paragraph missing");

    expect(property(block.paragraphFormatting.authored, "alignment")).toBe("right");
    expect(property(block.paragraphFormatting.effective, "alignment")).toBe("right");
    expect(property(block.runs[0]!.authoredFormatting, "italic")).toBe(true);
    expect(property(block.runs[0]!.effectiveFormatting, "bold")).toBe(true);
    expect(property(block.runs[0]!.effectiveFormatting, "italic")).toBe(true);
    expect(property(block.runs[0]!.effectiveFormatting, "fontSize")).toBe(22);
    expect(property(block.runs[0]!.effectiveFormatting, "strike")).toBeUndefined();

    paragraph.formatting!.alignment = "left";
    const run = paragraph.content.at(0);
    if (run?.type !== "run") throw new Error("authored run missing");
    run.formatting!.italic = false;
    const text = run.content.at(0);
    if (text?.type !== "text") throw new Error("authored text missing");
    text.text = "Mutated";
    paragraph.paraId = "A1000099";
    expect(property(block.paragraphFormatting.authored, "alignment")).toBe("right");
    expect(property(block.runs[0]!.authoredFormatting, "italic")).toBe(true);
    expect(block.identity.id).toBe("A1000001");
    expect(block.text).toBe("Alpha");
    expect(Object.isFrozen(block)).toBe(true);
    expect(Object.isFrozen(block.runs[0]?.authoredFormatting)).toBe(true);

    const operationSnapshot = resolvedDocxOperationSnapshot(snapshot);
    expect(operationSnapshot.blocks.at(0)?.text).toBe("Alpha");
    expect(sourceDocumentOf(operationSnapshot)).not.toBe(source);
    expect(sourceDocumentOf(operationSnapshot).eq(source)).toBe(true);
    expect(Object.isFrozen(operationSnapshot)).toBe(true);
    expect(Object.isFrozen(operationSnapshot.blocks)).toBe(true);
    expect(Object.isFrozen(operationSnapshot.blocks.at(0))).toBe(true);
    expect(Object.isFrozen(operationSnapshot.anchors)).toBe(true);
    expect(Object.isFrozen(numberingReferenceKeysOf(operationSnapshot))).toBe(true);
    expect(Object.isFrozen(storyTablesOf(operationSnapshot))).toBe(true);
  });

  test("rejects an equal-id source document whose text or markup differs", () => {
    const { document } = styledDocument();
    const { document: differentText, paragraph: textParagraph } = styledDocument();
    const textRun = textParagraph.content.at(0);
    if (textRun?.type !== "run") throw new Error("text fixture run missing");
    const text = textRun.content.at(0);
    if (text?.type !== "text") throw new Error("text fixture text missing");
    text.text = "Bravo";

    expect(() =>
      createResolvedDocxStorySnapshot({
        document,
        story: { type: "main" },
        sourceDocument: toProseDoc(differentText),
      }),
    ).toThrow("package projection disagree on source identity");

    const { document: differentMarkup, paragraph: markupParagraph } = styledDocument();
    const markupRun = markupParagraph.content.at(0);
    if (markupRun?.type !== "run") throw new Error("markup fixture run missing");
    markupRun.formatting = { ...markupRun.formatting, underline: { style: "single" } };
    expect(() =>
      createResolvedDocxStorySnapshot({
        document,
        story: { type: "main" },
        sourceDocument: toProseDoc(differentMarkup),
      }),
    ).toThrow("package projection disagree on source identity");
  });

  test("retains the caller's exact source identity for stale-state refusal", () => {
    const { document } = styledDocument();
    const sourceDocument = toProseDoc(document);
    const snapshot = createResolvedDocxStorySnapshot({
      document,
      story: { type: "main" },
      sourceDocument,
    });
    if (!snapshot) throw new Error("main story projection missing");

    expect(resolvedDocxSourceDocument(snapshot)).toBe(sourceDocument);
  });

  test("rejects same-id blocks from another capture at authored-run boundaries", () => {
    const capture = () => {
      const { document } = styledDocument();
      const snapshot = createResolvedDocxStorySnapshot({
        document,
        story: { type: "main" },
        sourceDocument: toProseDoc(document),
      });
      if (!snapshot) throw new Error("main story projection missing");
      const block = resolvedDocxContentBlocks(snapshot).at(0);
      if (!block) throw new Error("projected paragraph missing");
      return { snapshot, block };
    };
    const left = capture();
    const right = capture();
    expect(left.block).not.toBe(right.block);

    expect(() => resolvedDocxAuthoredRunsForBlock(left.snapshot, right.block)).toThrow(
      "must name its capsule's canonical block",
    );
    expect(() => resolvedDocxHasExactAuthoredRuns(left.snapshot, right.block)).toThrow(
      "must name its capsule's canonical block",
    );
  });

  test("projects container ownership by topology rather than package-local ids", () => {
    const project = (containerId: number) => {
      const { document, paragraph } = styledDocument();
      document.package.document.content = [
        {
          type: "blockSdt",
          properties: { sdtType: "richText", id: containerId },
          content: [paragraph],
        },
      ];
      const sourceDocument = toProseDoc(document);
      const snapshot = createResolvedDocxStorySnapshot({
        document,
        story: { type: "main" },
        sourceDocument,
      });
      if (!snapshot) throw new Error("main story projection missing");
      const block = resolvedDocxContentBlocks(snapshot).at(0);
      if (!block) throw new Error("projected paragraph missing");
      return block;
    };

    const first = project(17);
    const second = project(9001);
    expect(first.containerPath).toEqual([
      { kind: "blockSdt", identity: { type: "positional", id: "0" } },
    ]);
    expect(first.containerPath).toEqual(second.containerPath);
    expect(first.blockProperties).toEqual(second.blockProperties);
  });

  test("projects a nested table through its exact parent cell", () => {
    const paragraph = (paraId: string, text: string): Paragraph => ({
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    });
    const projectNested = (parentCellIndex: 0 | 1) => {
      const tableCell = (content: TableCell["content"]): TableCell => ({
        type: "tableCell",
        content,
      });
      const nested = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [paragraph("A1000010", "Nested")],
              },
            ],
          },
        ],
      } satisfies Table;
      const outer = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [0, 1].map((cellIndex) =>
              tableCell(
                cellIndex === parentCellIndex
                  ? [paragraph(`A100000${String(cellIndex + 1)}`, "Before"), nested]
                  : [paragraph(`A100000${String(cellIndex + 1)}`, "Other")],
              ),
            ),
          },
        ],
      } satisfies Table;
      const document = createEmptyDocument();
      document.package.document.content = [outer];
      const snapshot = createResolvedDocxStorySnapshot({
        document,
        story: { type: "main" },
        sourceDocument: toProseDoc(document),
      });
      if (!snapshot) throw new Error("main story projection missing");
      return (
        resolvedDocxContentBlocks(snapshot).find(({ text }) => text === "Nested") ??
        (() => {
          throw new Error("nested paragraph projection missing");
        })()
      );
    };

    expect(projectNested(0).containerPath).toEqual([
      {
        kind: "tableCell",
        identity: { type: "positional", id: "table-0-row-0-cell-0" },
      },
    ]);
    expect(projectNested(1).containerPath).toEqual([
      {
        kind: "tableCell",
        identity: { type: "positional", id: "table-0-row-0-cell-1" },
      },
    ]);
  });

  test("derives merged-cell geometry and paragraph ordinals from the live table", () => {
    const paragraph = (paraId: string, text: string): Paragraph => ({
      type: "paragraph",
      paraId,
      content: [
        {
          type: "run",
          content: text.length === 0 ? [] : [{ type: "text", text }],
        },
      ],
    });
    const table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              formatting: { vMerge: "restart" },
              content: [paragraph("A1000001", "First"), paragraph("A1000002", "Second")],
            },
            {
              type: "tableCell",
              content: [paragraph("A1000003", "Right")],
            },
          ],
        },
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              formatting: { vMerge: "continue" },
              content: [paragraph("A1000004", "")],
            },
            {
              type: "tableCell",
              content: [paragraph("A1000005", "Lower right")],
            },
          ],
        },
      ],
    } satisfies Table;
    const document = createEmptyDocument();
    document.package.document.content = [table];
    const sourceDocument = toProseDoc(document);
    const snapshot = createResolvedDocxStorySnapshot({
      document,
      story: { type: "main" },
      sourceDocument,
    });
    if (!snapshot) throw new Error("main story projection missing");

    expect(
      resolvedDocxContentBlocks(snapshot).map(({ identity, table: location }) => ({
        id: identity.id,
        row: location?.rowIndex,
        cell: location?.cellIndex,
        column: location?.gridColumnIndex,
        rowSpan: location?.rowSpan,
        paragraph: location?.paragraphIndex,
      })),
    ).toEqual([
      { id: "A1000001", row: 0, cell: 0, column: 0, rowSpan: 2, paragraph: 0 },
      { id: "A1000002", row: 0, cell: 0, column: 0, rowSpan: 2, paragraph: 1 },
      { id: "A1000003", row: 0, cell: 1, column: 1, rowSpan: 1, paragraph: 0 },
      { id: "A1000005", row: 1, cell: 0, column: 1, rowSpan: 1, paragraph: 0 },
    ]);
    const exposedTables = resolvedDocxTableNodes(snapshot);
    if (!(exposedTables instanceof Map)) throw new Error("table projection is not map-backed");
    exposedTables.clear();
    expect(resolvedDocxTableNodes(snapshot).size).toBe(1);
    expect(Object.isFrozen(storyTablesOf(resolvedDocxOperationSnapshot(snapshot)).at(0))).toBe(
      true,
    );
  });
});
