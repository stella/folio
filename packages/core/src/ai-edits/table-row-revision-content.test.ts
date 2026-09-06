/**
 * A tracked table-row insertion or deletion is marked twice: on the row
 * (`w:trPr/w:ins` | `w:trPr/w:del`) AND around every run in its cells
 * (`w:ins` | `w:del`, the latter carrying `w:delText`). Word writes both, and a
 * consumer that reads only run-level revisions — which is most of them — keeps
 * a deleted row's text on accept and an inserted row's text on reject when the
 * row marker stands alone.
 *
 * The fixtures are built from the typed model rather than authored in Word, so
 * the shape under test is spelled out in the test rather than hidden in bytes.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Table, TableCell, TableRow, TrackedChangeInfo } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const REVISION: TrackedChangeInfo = {
  id: 41,
  author: "Reviewer",
  date: "2026-01-02T03:04:05Z",
};

type RowSpec = {
  texts: readonly string[];
  revision?: "insertion" | "deletion";
};

const cell = (text: string, revision: RowSpec["revision"]): TableCell => {
  const run = { type: "run", content: [{ type: "text", text }] } as const;
  return {
    type: "tableCell",
    content: [
      {
        type: "paragraph",
        content: revision ? [{ type: revision, info: REVISION, content: [run] }] : [run],
      },
    ],
  };
};

const row = ({ texts, revision }: RowSpec): TableRow => ({
  type: "tableRow",
  cells: texts.map((text) => cell(text, revision)),
  ...(revision && {
    structuralChange: {
      type: revision === "insertion" ? "tableRowInsertion" : "tableRowDeletion",
      info: REVISION,
    },
  }),
});

const buildTableDocx = (rows: readonly RowSpec[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  const table: Table = { type: "table", rows: rows.map(row) };
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [table] },
    },
  });
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!entry) {
    throw new Error("expected word/document.xml");
  }
  return entry.async("string");
};

const tableRows = async (buffer: ArrayBuffer): Promise<TableRow[]> => {
  const parsed = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  const table = parsed.package.document.content.at(0);
  if (table?.type !== "table") {
    throw new Error("expected a table");
  }
  return table.rows;
};

const rowTexts = (rows: readonly TableRow[]): string[][] =>
  rows.map((tableRow) =>
    tableRow.cells.map((tableCell) =>
      tableCell.content
        .map((block) => (block.type === "paragraph" ? blockText(block) : ""))
        .join(""),
    ),
  );

const blockText = (paragraph: { content?: unknown[] }): string => {
  let text = "";
  const visit = (items: readonly unknown[]): void => {
    for (const item of items) {
      if (typeof item !== "object" || item === null) {
        continue;
      }
      const node = item as { type?: string; text?: string; content?: unknown[] };
      if (node.type === "text" && typeof node.text === "string") {
        text += node.text;
        continue;
      }
      if (Array.isArray(node.content)) {
        visit(node.content);
      }
    }
  };
  visit(paragraph.content ?? []);
  return text;
};

describe("a Word-style tracked row carries its cell content with it", () => {
  test("an inserted row accepts to the row and rejects to nothing", async () => {
    const buffer = await buildTableDocx([
      { texts: ["Kept"] },
      { texts: ["Added"], revision: "insertion" },
    ]);

    const accepting = await FolioDocxReviewer.fromBuffer(buffer);
    accepting.acceptAll();
    const accepted = await accepting.toBuffer();
    expect(rowTexts(await tableRows(accepted))).toEqual([["Kept"], ["Added"]]);
    expect(await documentXml(accepted)).not.toContain("<w:ins ");

    const rejecting = await FolioDocxReviewer.fromBuffer(buffer);
    rejecting.rejectAll();
    const rejected = await rejecting.toBuffer();
    expect(rowTexts(await tableRows(rejected))).toEqual([["Kept"]]);
    expect(await documentXml(rejected)).not.toContain("<w:ins ");
  });

  test("a deleted row accepts to nothing and rejects to the row", async () => {
    const buffer = await buildTableDocx([
      { texts: ["Kept"] },
      { texts: ["Removed"], revision: "deletion" },
    ]);

    const accepting = await FolioDocxReviewer.fromBuffer(buffer);
    accepting.acceptAll();
    const accepted = await accepting.toBuffer();
    expect(rowTexts(await tableRows(accepted))).toEqual([["Kept"]]);
    expect(await documentXml(accepted)).not.toContain("<w:del ");

    const rejecting = await FolioDocxReviewer.fromBuffer(buffer);
    rejecting.rejectAll();
    const rejected = await rejecting.toBuffer();
    expect(rowTexts(await tableRows(rejected))).toEqual([["Kept"], ["Removed"]]);
    const rejectedXml = await documentXml(rejected);
    expect(rejectedXml).not.toContain("<w:del ");
    expect(rejectedXml).not.toContain("<w:delText");
  });

  test("the row marker and its run marks are one change, not two", async () => {
    for (const revision of ["insertion", "deletion"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx([{ texts: ["Kept"] }, { texts: ["Moved"], revision }]),
      );
      expect(reviewer.getChanges().map(({ type }) => type)).toEqual([
        revision === "insertion" ? "rowInserted" : "rowDeleted",
      ]);
    }
  });

  test("serializing a tracked row is a byte fixed point", async () => {
    for (const revision of ["insertion", "deletion"] as const) {
      const source = await buildTableDocx([{ texts: ["Kept"] }, { texts: ["Marked"], revision }]);
      const once = await (await FolioDocxReviewer.fromBuffer(source)).toBuffer();
      const twice = await (await FolioDocxReviewer.fromBuffer(once)).toBuffer();
      expect(await documentXml(twice)).toBe(await documentXml(once));
    }
  });
});

describe("folio writes both halves of a row revision", () => {
  const findRowBlock = async (reviewer: FolioDocxReviewer, text: string): Promise<string> => {
    const block = reviewer.snapshot().blocks.find((candidate) => candidate.text === text);
    if (!block) {
      throw new Error(`expected a block reading "${text}"`);
    }
    return block.id;
  };

  test("a tracked row deletion marks the row and strikes its runs", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await buildTableDocx([{ texts: ["Alpha"] }, { texts: ["Beta"] }]),
      { author: "Reviewer" },
    );
    reviewer.applyOperations([
      { id: "delete-row", type: "deleteTableRow", blockId: await findRowBlock(reviewer, "Beta") },
    ]);

    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml).toContain("<w:trPr>");
    expect(xml).toContain("<w:delText");
    expect(xml).toContain("<w:del ");
  });

  test("a tracked row insertion marks the row and its runs", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await buildTableDocx([{ texts: ["Alpha"] }]),
      { author: "Reviewer" },
    );
    reviewer.applyOperations([
      {
        id: "insert-row",
        type: "insertTableRow",
        blockId: await findRowBlock(reviewer, "Alpha"),
        cellTexts: ["Gamma"],
      },
    ]);

    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml).toContain("<w:ins ");
    // The inserted run is inside the row, not only on the row marker.
    expect(/<w:ins [^>]*>\s*<w:r>/u.test(xml)).toBe(true);
  });

  test("resolving one revision clears the row marker and the run marks together", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await buildTableDocx([{ texts: ["Alpha"] }, { texts: ["Beta"] }]),
      { author: "Reviewer" },
    );
    reviewer.applyOperations([
      { id: "delete-row", type: "deleteTableRow", blockId: await findRowBlock(reviewer, "Beta") },
    ]);
    const change = reviewer.getChanges().at(0);
    if (!change) {
      throw new Error("expected a pending row deletion");
    }

    expect(reviewer.rejectChange(change)).toBe(true);
    expect(reviewer.getChanges()).toEqual([]);
    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml).not.toContain("<w:del ");
    expect(xml).not.toContain("<w:delText");
  });
});

describe("row insert/delete round trips through the reviewed views", () => {
  const cellWord = fc.stringMatching(/^[A-Za-z]{2,8}$/u);

  /**
   * The reviewed view prefixes each block with its stable id, which is derived
   * from content and position and so differs between two reviewers holding the
   * same text. The comparison is about the text the views resolve to.
   */
  const storyText = (reviewer: FolioDocxReviewer, view: "original" | "final"): string => {
    const story = reviewer.readReviewedStory({ view });
    if (!story) {
      throw new Error("expected the main story");
    }
    return story.text.replaceAll(/^\[[^\]]+\]\s?/gmu, "");
  };

  test(
    "the final view is the target and the original view is the base",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.array(fc.array(cellWord, { minLength: 1, maxLength: 3 }), {
            minLength: 1,
            maxLength: 4,
          }),
          fc.constantFrom("insertTableRow" as const, "deleteTableRow" as const),
          fc.nat(),
          fc.array(cellWord, { minLength: 1, maxLength: 3 }),
          async (grid, type, selector, cellTexts) => {
            // A ragged grid is not a table any consumer would write; pad every
            // row to the widest one so the fixture is a real rectangle.
            const width = Math.max(...grid.map((cells) => cells.length));
            const rows = grid.map((cells) => ({
              texts: [...cells, ...Array.from({ length: width - cells.length }, () => "pad")],
            }));
            const buffer = await buildTableDocx(rows);

            const base = await FolioDocxReviewer.fromBuffer(buffer, { author: "Reviewer" });
            const blocks = base.snapshot().blocks;
            if (blocks.length === 0) {
              return;
            }
            const blockId = blocks[selector % blocks.length]!.id;
            const operation =
              type === "insertTableRow"
                ? { id: "row", type, blockId, cellTexts }
                : { id: "row", type, blockId };

            const direct = await FolioDocxReviewer.fromBuffer(buffer, { author: "Reviewer" });
            if (direct.applyOperations([operation], { mode: "direct" }).applied.length === 0) {
              return;
            }

            const tracked = await FolioDocxReviewer.fromBuffer(buffer, { author: "Reviewer" });
            if (tracked.applyOperations([operation]).applied.length === 0) {
              return;
            }

            expect(storyText(tracked, "final")).toBe(storyText(direct, "final"));
            expect(storyText(tracked, "original")).toBe(storyText(base, "final"));
          },
        ),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
