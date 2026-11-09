import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { reconcileRenderedContent } from "./reconcileRenderedContent";

setDefaultTimeout(propertyTestTimeout(10_000));

type Line = {
  pmStart: number;
  text: string;
  bold: boolean;
  spans: number;
};

/** A painted paragraph: a block with its position and some spans of text. */
const paragraphElement = (line: Line, shift: number): HTMLElement => {
  const paragraph = document.createElement("div");
  paragraph.className = "layout-paragraph";
  paragraph.dataset["pmStart"] = String(line.pmStart + shift);
  paragraph.dataset["pmEnd"] = String(line.pmStart + shift + line.text.length);
  for (let index = 0; index < line.spans; index += 1) {
    const span = document.createElement("span");
    span.dataset["pmStart"] = String(line.pmStart + shift + index);
    if (line.bold) span.style.setProperty("font-weight", "700");
    span.textContent = `${line.text}${index}`;
    paragraph.append(span);
  }
  return paragraph;
};

const contentElement = (lines: readonly Line[], shifts: readonly number[]) => {
  const content = document.createElement("div");
  content.className = "layout-page-content";
  for (const [index, line] of lines.entries())
    content.append(paragraphElement(line, shifts[index] ?? 0));
  return content;
};

const line = fc.record({
  pmStart: fc.integer({ min: 0, max: 500 }),
  text: fc.constantFrom("alpha", "beta", "gamma", "x"),
  bold: fc.boolean(),
  spans: fc.integer({ min: 0, max: 3 }),
});

/** What an edit does to one painted line in the fresh rendering. */
const change = fc.constantFrom(
  "same",
  "shifted",
  "text",
  "bold",
  "spans",
) satisfies fc.Arbitrary<"same" | "shifted" | "text" | "bold" | "spans">;

test("reconciling painted content gives the fresh rendering and keeps the lines it shows the same", () => {
  const releaseDom = acquireHarnessDom();
  try {
    assertProperty(
      fc.property(
        fc.array(fc.tuple(line, change), { minLength: 0, maxLength: 8 }),
        fc.integer({ min: -3, max: 3 }),
        fc.boolean(),
        (lines, delta, grows) => {
          const before = lines.map(([painted]) => painted);
          const after = lines.map(([painted, kind]): Line => {
            switch (kind) {
              case "same":
              case "shifted":
                return painted;
              case "text":
                return { ...painted, text: `${painted.text}!` };
              case "bold":
                return { ...painted, bold: !painted.bold };
              case "spans":
                return { ...painted, spans: painted.spans + 1 };
              default: {
                const exhaustive: never = kind;
                return exhaustive;
              }
            }
          });
          const shifts = lines.map(([, kind]) => (kind === "same" ? 0 : delta));
          if (grows)
            after.push({ pmStart: 0, text: "new", bold: false, spans: 1 });
          const painted = contentElement(
            before,
            before.map(() => 0),
          );
          const fresh = contentElement(after, shifts);
          const expected = fresh.cloneNode(true);
          const kept = Array.from(painted.childNodes);
          const reconciled = reconcileRenderedContent(painted, fresh);
          if (grows) {
            // A different number of lines is the caller's to replace whole.
            expect(reconciled).toBe(false);
            return;
          }
          expect(reconciled).toBe(true);
          // The content is now exactly what painting it afresh gives.
          expect(painted.isEqualNode(expected)).toBe(true);
          // And a line the edit only shifted, or left alone, is the node that
          // was painted before.
          for (const [index, [, kind]] of lines.entries()) {
            const reused = painted.childNodes[index] === kept[index];
            if (kind === "same" || kind === "shifted")
              expect(reused).toBe(true);
          }
        },
      ),
      {
        id: "reconciling painted content gives the fresh rendering and keeps the lines it shows the same",
      },
    );
  } finally {
    releaseDom();
  }
});

type Edit = "same" | "shifted" | "text" | "bold" | "spans";

const edited = (painted: Line, kind: Edit): Line => {
  switch (kind) {
    case "same":
    case "shifted":
      return painted;
    case "text":
      return { ...painted, text: `${painted.text}!` };
    case "bold":
      return { ...painted, bold: !painted.bold };
    case "spans":
      return { ...painted, spans: painted.spans + 1 };
    default: {
      const exhaustive: never = kind;
      return exhaustive;
    }
  }
};

/** A table's cells, each holding paragraphs, with an edit for each paragraph. */
type Table = {
  rows: (readonly (readonly [Line, Edit])[])[][];
  /** What the edit does to the table itself. */
  shell: "same" | "width" | "row";
};

type Block =
  | { kind: "paragraph"; line: Line; edit: Edit }
  | { kind: "table"; table: Table };

const tableElement = (
  rows: readonly (readonly Line[])[][],
  shifts: readonly (readonly number[])[][],
  width: number,
): HTMLElement => {
  const table = document.createElement("div");
  table.className = "layout-table";
  table.style.setProperty("width", `${String(width)}px`);
  for (const [rowIndex, cells] of rows.entries()) {
    const row = document.createElement("div");
    row.className = "layout-table-row";
    for (const [cellIndex, lines] of cells.entries()) {
      const cell = document.createElement("div");
      cell.className = "layout-table-cell";
      const content = document.createElement("div");
      content.className = "layout-table-cell-content";
      for (const [index, paragraph] of lines.entries())
        content.append(
          paragraphElement(
            paragraph,
            shifts[rowIndex]?.[cellIndex]?.[index] ?? 0,
          ),
        );
      cell.append(content);
      row.append(cell);
    }
    table.append(row);
  }
  return table;
};

const edit = change as fc.Arbitrary<Edit>;
const table: fc.Arbitrary<Table> = fc.record({
  rows: fc.array(
    fc.array(fc.array(fc.tuple(line, edit), { maxLength: 3 }), {
      minLength: 1,
      maxLength: 3,
    }),
    { minLength: 1, maxLength: 3 },
  ),
  shell: fc.constantFrom("same", "same", "same", "width", "row"),
});
const block: fc.Arbitrary<Block> = fc.oneof(
  fc.record({
    kind: fc.constant("paragraph" as const),
    line,
    edit,
  }),
  fc.record({ kind: fc.constant("table" as const), table }),
);

test("reconciling nested content replaces only the smallest parts that changed", () => {
  const releaseDom = acquireHarnessDom();
  try {
    assertProperty(
      fc.property(
        fc.array(block, { minLength: 0, maxLength: 5 }),
        fc.integer({ min: -3, max: 3 }),
        (blocks, delta) => {
          const shiftOf = (kind: Edit) => (kind === "same" ? 0 : delta);
          const render = (fresh: boolean) => {
            const content = document.createElement("div");
            content.className = "layout-page-content";
            for (const item of blocks) {
              if (item.kind === "paragraph") {
                content.append(
                  paragraphElement(
                    fresh ? edited(item.line, item.edit) : item.line,
                    fresh ? shiftOf(item.edit) : 0,
                  ),
                );
                continue;
              }
              const { rows, shell } = item.table;
              const lines = rows.map((cells) =>
                cells.map((cell) =>
                  cell.map(([painted, kind]) =>
                    fresh ? edited(painted, kind) : painted,
                  ),
                ),
              );
              const shifts = rows.map((cells) =>
                cells.map((cell) =>
                  cell.map(([, kind]) => (fresh ? shiftOf(kind) : 0)),
                ),
              );
              if (fresh && shell === "row") lines.push([[]]);
              content.append(
                tableElement(
                  lines,
                  shifts,
                  fresh && shell === "width" ? 101 : 100,
                ),
              );
            }
            return content;
          };
          const painted = render(false);
          const fresh = render(true);
          const expected = fresh.cloneNode(true);
          const kept = Array.from(painted.childNodes);
          const paragraphsOf = (element: Element | undefined) =>
            Array.from(element?.querySelectorAll(".layout-paragraph") ?? []);
          const keptParagraphs = kept.map((node) =>
            paragraphsOf(node as Element),
          );

          expect(reconcileRenderedContent(painted, fresh)).toBe(true);
          // The content is now exactly what painting it afresh gives.
          expect(painted.isEqualNode(expected)).toBe(true);
          for (const [index, item] of blocks.entries()) {
            const node = painted.childNodes[index];
            if (item.kind === "paragraph") {
              if (item.edit === "same" || item.edit === "shifted")
                expect(node).toBe(kept[index]);
              continue;
            }
            const { rows, shell } = item.table;
            // A table whose own attributes and rows stay is kept, and so is
            // every paragraph in it the edit only shifted or left alone.
            if (shell !== "same") continue;
            expect(node).toBe(kept[index]);
            const now = paragraphsOf(node as Element);
            const before = keptParagraphs[index] ?? [];
            const edits = rows.flat(2).map(([, kind]) => kind);
            for (const [position, kind] of edits.entries())
              if (kind === "same" || kind === "shifted")
                expect(now[position]).toBe(before[position]);
          }
        },
      ),
      {
        id: "reconciling nested content replaces only the smallest parts that changed",
      },
    );
  } finally {
    releaseDom();
  }
});
