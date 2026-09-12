import { describe, expect, test } from "bun:test";

import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import type { FolioContentPropertySet } from "../../compare/content-types";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Paragraph, Table } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxAuthoredRunsForBlock,
  resolvedDocxContentBlocks,
  resolvedDocxHasExactAuthoredRuns,
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
    const operationSnapshot = createFolioAIEditSnapshot(source);
    Reflect.set(operationSnapshot.blocks[0]!, "previewRuns", [
      { text: "Alpha", effectiveFormatting: { strike: true } },
    ]);
    const snapshot = createResolvedDocxStorySnapshot({
      document,
      story: { type: "main" },
      operationSnapshot,
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
    expect(property(block.paragraphFormatting.authored, "alignment")).toBe("right");
    expect(property(block.runs[0]!.authoredFormatting, "italic")).toBe(true);
    expect(Object.isFrozen(block)).toBe(true);
    expect(Object.isFrozen(block.runs[0]?.authoredFormatting)).toBe(true);
  });

  test("types a non-lossless authored text projection as unsupported", () => {
    const { document } = styledDocument();
    const source = toProseDoc(document);
    const operationSnapshot = createFolioAIEditSnapshot(source);
    Reflect.set(operationSnapshot.blocks[0]!, "text", "Bravo");
    const snapshot = createResolvedDocxStorySnapshot({
      document,
      story: { type: "main" },
      operationSnapshot,
    });
    if (!snapshot) throw new Error("main story projection missing");
    const block = resolvedDocxContentBlocks(snapshot).at(0);
    if (!block) throw new Error("projected paragraph missing");

    expect(block.text).toBe("Bravo");
    expect(resolvedDocxHasExactAuthoredRuns(snapshot, block)).toBe(false);
    expect(() => resolvedDocxAuthoredRunsForBlock(snapshot, block)).toThrow(
      "lossy DOCX authored-run projection reached transport lowering",
    );
    expect(property(block.blockProperties, "docx.authoredRunProjection")).toBeDefined();
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
      const operationSnapshot = createFolioAIEditSnapshot(toProseDoc(document));
      const snapshot = createResolvedDocxStorySnapshot({
        document,
        story: { type: "main" },
        operationSnapshot,
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
    const operationSnapshot = createFolioAIEditSnapshot(toProseDoc(document));
    const snapshot = createResolvedDocxStorySnapshot({
      document,
      story: { type: "main" },
      operationSnapshot,
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
  });
});
