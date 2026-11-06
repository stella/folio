/**
 * What a row omits, stated by `w:gridBefore` and `w:gridAfter` or by stopping
 * short, is part of where its cells stand. Row and column edits work on such a
 * table, direct and tracked agree, a rejection restores the source, and no
 * cell's text changes column.
 */

import { describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import { CellSelection, TableMap } from "prosemirror-tables";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createHarnessState,
  harnessManager,
  parseShapeDocument,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

/** A row: cells by text, what it omits before and states it omits after. */
type RowShape = { cells: string[]; before?: number; after?: number };

const SHAPES: Record<string, { width: number; rows: RowShape[] }> = {
  "rows stopping short": {
    width: 3,
    rows: [{ cells: ["a", "b", "c"] }, { cells: ["d"] }, { cells: ["e", "f"] }],
  },
  "a declared w:gridAfter": {
    width: 3,
    rows: [
      { cells: ["a", "b", "c"] },
      { cells: ["d"], after: 2 },
      { cells: ["e", "f"], after: 1 },
    ],
  },
  "a declared w:gridBefore": {
    width: 3,
    rows: [
      { cells: ["a", "b", "c"] },
      { cells: ["d"], before: 2 },
      { cells: ["e", "f"], before: 1 },
    ],
  },
  "both at once": {
    width: 4,
    rows: [
      { cells: ["a", "b", "c", "g"] },
      { cells: ["d", "e"], before: 1, after: 1 },
      { cells: ["f"] },
    ],
  },
};

const idOf = (text: string) =>
  `1000${text.charCodeAt(0).toString(16).padStart(4, "0")}`;

const documentXml = (width: number, rows: readonly RowShape[]) =>
  `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(width)}</w:tblGrid>${rows
    .map(({ cells, before, after }) => {
      const properties = `${before ? `<w:gridBefore w:val="${before}"/>` : ""}${after ? `<w:gridAfter w:val="${after}"/>` : ""}`;
      return `<w:tr>${properties ? `<w:trPr>${properties}</w:trPr>` : ""}${cells
        .map(
          (text) =>
            `<w:tc><w:tcPr/><w:p w14:paraId="${idOf(text)}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`,
        )
        .join("")}</w:tr>`;
    })
    .join(
      "",
    )}</w:tbl><w:p w14:paraId="10000001"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`;

const docxOf = async (width: number, rows: readonly RowShape[]) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set("word/document.xml", documentXml(width, rows));
  return (await zip.generate()).slice().buffer;
};

/** Each saved row as its cells' text by the grid column they start at. */
const savedColumns = async (buffer: ArrayBuffer) => {
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [...xml.matchAll(/<w:tr[ >].*?<\/w:tr>/gsu)].map(([row]) => {
    let column = Number(/<w:gridBefore w:val="(\d+)"/u.exec(row)?.[1] ?? 0);
    const cells: Record<string, number> = {};
    for (const [cell] of row.matchAll(/<w:tc>.*?<\/w:tc>/gsu)) {
      const text = [...cell.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
        .map(([, content]) => content)
        .join("");
      if (text) cells[text] = column;
      column += Number(/<w:gridSpan w:val="(\d+)"/u.exec(cell)?.[1] ?? 1);
    }
    return cells;
  });
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0));

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operation: Record<string, unknown>,
) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [{ id: "edit", ...operation }] } as never,
    { undefinedReferences: "refuse" },
  );

const flat = (rows: readonly Record<string, number>[]) =>
  Object.assign({}, ...rows) as Record<string, number>;

const columnOf = (rows: readonly RowShape[], text: string) => {
  for (const { cells, before } of rows) {
    const at = cells.indexOf(text);
    if (at >= 0) return at + (before ?? 0);
  }
  throw new Error(`No cell ${text}`);
};

describe.each(Object.entries(SHAPES))("a table with %s", (_name, shape) => {
  const { width, rows } = shape;
  const texts = rows.flatMap(({ cells }) => cells);
  const load = () => docxOf(width, rows);

  test("opens with a rectangular map and saves as it was", async () => {
    const buffer = await load();
    const state = createHarnessState(
      await parseShapeDocument(new Uint8Array(buffer.slice(0))),
      "editing",
    );
    const map = TableMap.get(state.doc.child(0));
    expect(map.problems).toBeNull();
    expect(map.width).toBe(width);
    const saved = await savedColumns(await (await open(buffer)).toBuffer());
    expect(flat(saved)).toEqual(
      Object.fromEntries(texts.map((text) => [text, columnOf(rows, text)])),
    );
  });

  describe.each(texts)("from cell %s", (anchor) => {
    const blockId = idOf(anchor);
    const sourceColumn = columnOf(rows, anchor);

    const expectAgreement = async (
      operation: Record<string, unknown>,
      expected: Record<string, number>,
    ) => {
      const buffer = await load();
      const source = await savedColumns(await (await open(buffer)).toBuffer());

      const direct = await open(buffer);
      expect(apply(direct, "direct", operation).applied).toHaveLength(1);
      const directSaved = await direct.toBuffer();
      const directColumns = await savedColumns(directSaved);
      expect(flat(directColumns)).toEqual(expected);

      const tracked = await open(buffer);
      expect(apply(tracked, "tracked-changes", operation).applied).toHaveLength(
        1,
      );
      const trackedSaved = await tracked.toBuffer();

      const accepted = await open(trackedSaved);
      accepted.acceptAll();
      expect(await savedColumns(await accepted.toBuffer())).toEqual(
        directColumns,
      );

      const live = await open(buffer);
      apply(live, "tracked-changes", operation);
      live.acceptAll();
      expect(await savedColumns(await live.toBuffer())).toEqual(directColumns);

      const rejected = await open(trackedSaved);
      rejected.rejectAll();
      expect(await savedColumns(await rejected.toBuffer())).toEqual(source);

      const liveRejected = await open(buffer);
      apply(liveRejected, "tracked-changes", operation);
      liveRejected.rejectAll();
      expect(await savedColumns(await liveRejected.toBuffer())).toEqual(source);
    };

    const unchanged = Object.fromEntries(
      texts.map((text) => [text, columnOf(rows, text)]),
    );

    test.each(["before", "after"] as const)(
      "takes a column %s it",
      async (position) => {
        const at = position === "before" ? sourceColumn : sourceColumn + 1;
        await expectAgreement(
          { type: "insertTableColumn", blockId, position },
          Object.fromEntries(
            texts.map((text) => {
              const column = columnOf(rows, text);
              return [text, column >= at ? column + 1 : column];
            }),
          ),
        );
      },
    );

    test("takes a row after it", async () => {
      await expectAgreement(
        {
          type: "insertTableRow",
          blockId,
          position: "after",
          cellTexts: ["N"],
        },
        { ...unchanged, N: 0 },
      );
    });

    test("loses its row", async () => {
      const row = rows.find(({ cells }) => cells.includes(anchor));
      await expectAgreement(
        { type: "deleteTableRow", blockId },
        Object.fromEntries(
          texts
            .filter((text) => !row?.cells.includes(text))
            .map((text) => [text, columnOf(rows, text)]),
        ),
      );
    });

    test("loses its column", async () => {
      await expectAgreement(
        { type: "deleteTableColumn", blockId },
        Object.fromEntries(
          texts
            .filter((text) => columnOf(rows, text) !== sourceColumn)
            .map((text) => {
              const column = columnOf(rows, text);
              return [text, column > sourceColumn ? column - 1 : column];
            }),
        ),
      );
    });
  });
});

describe("merging across what a row omits", () => {
  test("is refused, in either mode", async () => {
    const buffer = await docxOf(3, SHAPES["rows stopping short"]?.rows ?? []);
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await open(buffer);
      const result = apply(reviewer, mode, {
        type: "mergeTableCells",
        blockId: idOf("c"),
        endBlockId: idOf("f"),
      });
      expect(result.applied).toHaveLength(0);
    }
  });

  test("beside it works", async () => {
    const buffer = await docxOf(3, SHAPES["rows stopping short"]?.rows ?? []);
    const reviewer = await open(buffer);
    const result = apply(reviewer, "direct", {
      type: "mergeTableCells",
      blockId: idOf("a"),
      endBlockId: idOf("b"),
    });
    expect(result.applied).toHaveLength(1);
    expect(flat(await savedColumns(await reviewer.toBuffer()))).toMatchObject({
      d: 0,
      e: 0,
      f: 1,
    });
  });
});

describe.each(Object.entries(SHAPES))(
  "the editor's commands on a table with %s",
  (_name, { width, rows }) => {
    const texts = rows.flatMap(({ cells }) => cells);

    const run = async (command: string, anchor: string) => {
      const base = await parseShapeDocument(
        new Uint8Array(await docxOf(width, rows)),
      );
      const state = createHarnessState(base, "editing");
      let at = -1;
      state.doc.descendants((node, position) => {
        if (node.isTextblock && node.textContent === anchor) at = position + 1;
        return at < 0;
      });
      const selected = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, at)),
      );
      const factory = harnessManager().getCommand(command);
      if (!factory) throw new Error(`Missing command: ${command}`);
      let next = selected;
      const handled = factory()(selected, (tr: Transaction) => {
        next = selected.apply(tr);
      });
      expect(handled).toBe(true);
      const table = next.doc.child(0);
      expect(TableMap.get(table).problems).toBeNull();
      const { bytes } = await saveHarnessState(next, base);
      return savedColumns(bytes.slice().buffer);
    };

    describe.each(texts)("from cell %s", (anchor) => {
      const sourceColumn = columnOf(rows, anchor);
      const shifted = (at: number) =>
        Object.fromEntries(
          texts.map((text) => {
            const column = columnOf(rows, text);
            return [text, column >= at ? column + 1 : column];
          }),
        );

      test("adds a column to its left and right", async () => {
        expect(flat(await run("addColumnLeft", anchor))).toEqual(
          shifted(sourceColumn),
        );
        expect(flat(await run("addColumnRight", anchor))).toEqual(
          shifted(sourceColumn + 1),
        );
      });

      test("adds rows above and below", async () => {
        const unchanged = shifted(width + 1);
        expect(flat(await run("addRowAbove", anchor))).toEqual(unchanged);
        expect(flat(await run("addRowBelow", anchor))).toEqual(unchanged);
      });

      test("deletes its row", async () => {
        const row = rows.find(({ cells }) => cells.includes(anchor));
        expect(flat(await run("deleteRow", anchor))).toEqual(
          Object.fromEntries(
            texts
              .filter((text) => !row?.cells.includes(text))
              .map((text) => [text, columnOf(rows, text)]),
          ),
        );
      });

      test("deletes its column", async () => {
        expect(flat(await run("deleteColumn", anchor))).toEqual(
          Object.fromEntries(
            texts
              .filter((text) => columnOf(rows, text) !== sourceColumn)
              .map((text) => {
                const column = columnOf(rows, text);
                return [text, column > sourceColumn ? column - 1 : column];
              }),
          ),
        );
      });
    });
  },
);

test("a grid wider than the rows keeps its widths through a column edit", async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    documentXml(2, [{ cells: ["a", "b"] }]).replace(
      /<w:tblGrid>.*<\/w:tblGrid>/u,
      '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>',
    ),
  );
  const reviewer = await open((await zip.generate()).slice().buffer);
  apply(reviewer, "direct", {
    type: "insertTableColumn",
    blockId: idOf("a"),
    position: "after",
  });
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  expect(
    [...xml.matchAll(/<w:gridCol w:w="(\d+)"/gu)].map(([, width]) =>
      Number(width),
    ),
  ).toEqual([1000, 1000, 2000, 3000]);
});

describe("the editor's merge and split on a table with short rows", () => {
  const rows = SHAPES["rows stopping short"]?.rows ?? [];

  const stateWithCells = async (from: string, to: string) => {
    const base = await parseShapeDocument(
      new Uint8Array(await docxOf(3, rows)),
    );
    const state = createHarnessState(base, "editing");
    const cellAt = (text: string) => {
      let found = -1;
      state.doc.descendants((node, position) => {
        if (
          node.type.spec["tableRole"] === "cell" &&
          node.textContent === text
        ) {
          found = position;
        }
        return found < 0;
      });
      return found;
    };
    const selected = state.apply(
      state.tr.setSelection(
        CellSelection.create(state.doc, cellAt(from), cellAt(to)),
      ),
    );
    return { base, selected };
  };

  const run = (
    selected: Awaited<ReturnType<typeof stateWithCells>>["selected"],
    command: string,
  ) => {
    const factory = harnessManager().getCommand(command);
    if (!factory) throw new Error(`Missing command: ${command}`);
    let next = selected;
    const handled = factory()(selected, (tr: Transaction) => {
      next = selected.apply(tr);
    });
    return { handled, next };
  };

  test("merges cells beside what a row omits, then splits them again", async () => {
    const { base, selected } = await stateWithCells("a", "b");
    const merged = run(selected, "mergeCells");
    expect(merged.handled).toBe(true);
    expect(TableMap.get(merged.next.doc.child(0)).problems).toBeNull();
    const { bytes } = await saveHarnessState(merged.next, base);
    expect(flat(await savedColumns(bytes.slice().buffer))).toEqual({
      ab: 0,
      c: 2,
      d: 0,
      e: 0,
      f: 1,
    });

    const split = run(merged.next, "splitCell");
    expect(split.handled).toBe(true);
    expect(TableMap.get(split.next.doc.child(0)).problems).toBeNull();
    const resplit = await saveHarnessState(split.next, base);
    expect(flat(await savedColumns(resplit.bytes.slice().buffer))).toEqual({
      ab: 0,
      c: 2,
      d: 0,
      e: 0,
      f: 1,
    });
  });

  test("refuses to merge through what a row omits", async () => {
    const { selected } = await stateWithCells("c", "f");
    expect(run(selected, "mergeCells").handled).toBe(false);
  });
});
